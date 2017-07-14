using System;
using System.Windows.Forms;
using System.Collections.Generic;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for Tables.
	/// </summary>
	public class Queries
	{
		public int m_intError=0;
		public string m_strError="";
		public Project m_oProject = new Project();
		public CoreScenarioResults m_oCoreScenarioResults = new CoreScenarioResults();
		public CoreScenarioRuleDefinitions m_oCoreScenarioRuleDef = new CoreScenarioRuleDefinitions();
        public Queries.CoreScenarioRun m_oCoreAnalysisScenarioRun = new CoreScenarioRun();
		public FIAPlot m_oFIAPlot = new FIAPlot();
		public FVS m_oFvs = new FVS();
		public TravelTime m_oTravelTime = new TravelTime();
		public Processor m_oProcessor = new Processor();
        public ProcessorScenarioRun m_oProcessorScenarioRun = new ProcessorScenarioRun();
		public Audit m_oAudit = new Audit();
		public Reference m_oReference = new Reference();
		public FIA_Biosum_Manager.Datasource m_oDataSource;
		public string m_strTempDbFile;
		private bool _bScenario=false;
		private string _strScenarioType="";

		
		public Queries()
		{
			m_oFvs.ReferenceQueries=this;
			m_oFIAPlot.ReferenceQueries=this;
			m_oReference.ReferenceQueries=this;
            m_oProcessor.ReferenceQueries = this;
            m_oTravelTime.ReferenceQueries = this;
			//
			// TODO: Add constructor logic here
			//
		}
		public bool Scenario
		{
			get {return _bScenario;}
			set {_bScenario=value;}
		}
		public string ScenarioType
		{
			get {return _strScenarioType;}
			set {_strScenarioType=value;}
		}
		/// <summary>
		/// use the DataSource class to get DB files and table names.
		/// </summary>
		/// <param name="p_bLimited">TRUE=do not open and load the table record numbers, table column names, and table column data types</param>
		public void LoadDatasources(bool  p_bLimited)
		{
			if (p_bLimited)
			{
				LoadLimitedDatasources();
				
			}
			else
			{
				
			}
			if (this.m_oFvs.LoadDatasource) this.m_oFvs.LoadDatasources();
			if (this.m_oFIAPlot.LoadDatasource) this.m_oFIAPlot.LoadDatasources();
			if (this.m_oReference.LoadDatasource) this.m_oReference.LoadDatasources();
            if (this.m_oProcessor.LoadDatasource) this.m_oProcessor.LoadDatasources();
            if (this.m_oTravelTime.LoadDatasource) this.m_oTravelTime.LoadDatasources();
			m_strTempDbFile = this.m_oDataSource.CreateMDBAndTableDataSourceLinks();
		}
		public void LoadDatasources(bool p_bLimited,string p_strScenarioType,string p_strScenarioId)
		{
			Scenario=true;
			ScenarioType=p_strScenarioType;
			if (p_bLimited)
			{
				LoadLimitedDatasources(p_strScenarioType,p_strScenarioId);
				
			}
			else
			{
				
			}
			if (this.m_oFvs.LoadDatasource) this.m_oFvs.LoadDatasources();
			if (this.m_oFIAPlot.LoadDatasource) this.m_oFIAPlot.LoadDatasources();
			if (this.m_oReference.LoadDatasource) this.m_oReference.LoadDatasources();
            if (this.m_oProcessor.LoadDatasource) this.m_oProcessor.LoadDatasources();
			m_strTempDbFile = this.m_oDataSource.CreateMDBAndTableDataSourceLinks();
		}

		protected void LoadLimitedDatasources()
		{
			string strProjDir=frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
			
			m_oDataSource = new Datasource();
			m_oDataSource.LoadTableColumnNamesAndDataTypes=false;
			m_oDataSource.LoadTableRecordCount=false;
			m_oDataSource.m_strDataSourceMDBFile = strProjDir.Trim() + "\\db\\project.mdb";
			m_oDataSource.m_strDataSourceTableName = "datasource";
			m_oDataSource.m_strScenarioId="";
			m_oDataSource.populate_datasource_array();
			
			
		}
		protected void LoadLimitedDatasources(string p_strScenarioType,string p_strScenarioId)
		{
			string strProjDir=frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
			
			m_oDataSource = new Datasource();
			m_oDataSource.LoadTableColumnNamesAndDataTypes=false;
			m_oDataSource.m_strScenarioId=p_strScenarioId.Trim();
			m_oDataSource.LoadTableRecordCount=false;
			m_oDataSource.m_strDataSourceMDBFile = strProjDir.Trim() + "\\" + p_strScenarioType + "\\db\\scenario_" + p_strScenarioType + "_rule_definitions.mdb";
			m_oDataSource.m_strDataSourceTableName = "scenario_datasource";
			m_oDataSource.populate_datasource_array();
			
		}
		static public string GetInsertSQL(string p_strFields, string p_strValues,string p_strTable)
		{
			return "INSERT INTO " + p_strTable + " (" + p_strFields + ") VALUES (" + p_strValues + ")";
		}
		public static string GenericSelectSQL(string p_strTable,string p_strFields,string p_strWhereExpression)
		{
			return "SELECT " + p_strFields + " FROM " + p_strTable + " WHERE " + p_strWhereExpression;
		}
		public class Project
		{
			
		}
		public class CoreScenarioResults
		{
			private string strSQL = "";
			public CoreScenarioResults()
			{
			}
		}
		public class CoreScenarioRuleDefinitions
		{
			private string strSQL = "";
			public CoreScenarioRuleDefinitions()
			{
			}
		}
        public class CoreScenarioRun
        {
            private string strSQL = "";
            private string _strScenarioId = "";
            public string ScenarioId
            {
                get { return _strScenarioId; }
                set { _strScenarioId = value; }
            }
            public CoreScenarioRun()
            {
            }
        }
		public class FVS
		{
			public string m_strRxTable;
			public string m_strFvsCmdTable;
			public string m_strFvsCatTable;
			public string m_strFvsSubCatTable;
			public string m_strRxFvsCmdTable;
			public string m_strRxHarvestCostColumnsTable;
			public string m_strRxPackageTable;
			public string m_strRxPackageMembersTable;
			public string m_strRxPackageFvsCmdTable;
			public string m_strRxPackageFvsCmdOrderTable;
			public string m_strTreeSpcTable;
			public string m_strFvsTreeTable;
			public string m_strFvsTreeSpcRefTable;
            public string m_strFVSWesternTreeSpeciesTable;
            public string m_strFVSEasternTreeSpeciesTable;
            public string m_strFVSPrePostSeqNumTable;
            public string m_strFVSPrePostSeqNumRxPackageAssgnTable;
			private Queries _oQueries=null;	
			private bool _bLoadDataSources=true;
			string m_strSql="";

			public FVS()
			{
			}
			public Queries ReferenceQueries
			{
				get {return _oQueries;}
				set {_oQueries=value;}
			}
			public bool LoadDatasource
			{
				get {return _bLoadDataSources;}
				set {_bLoadDataSources=value;}
			}


			public void LoadDatasources()
			{
				m_strRxTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PRESCRIPTIONS");
				m_strRxFvsCmdTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PRESCRIPTIONS ASSIGNED FVS COMMANDS");
				m_strRxHarvestCostColumnsTable=ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PRESCRIPTIONS HARVEST COST COLUMNS");
				m_strFvsCmdTable  = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("FVS COMMANDS");
                m_strFVSWesternTreeSpeciesTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("FVS WESTERN TREE SPECIES TRANSLATOR");
                m_strFVSEasternTreeSpeciesTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("FVS EASTERN TREE SPECIES TRANSLATOR");
				m_strFvsCatTable =  ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PRESCRIPTION CATEGORIES");
				m_strFvsSubCatTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PRESCRIPTION SUBCATEGORIES");
				m_strRxPackageTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PACKAGES");
				m_strRxPackageMembersTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PACKAGE MEMBERS");
				m_strRxPackageFvsCmdTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PACKAGE ASSIGNED FVS COMMANDS");
				m_strRxPackageFvsCmdOrderTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PACKAGE FVS COMMANDS ORDER");
				m_strTreeSpcTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREE SPECIES");
				m_strFvsTreeTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("FVS TREE LIST FOR PROCESSOR");
				m_strFvsTreeSpcRefTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("FVS TREE SPECIES");
                m_strFVSPrePostSeqNumTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("FVS PRE-POST SEQNUM DEFINITIONS");
                m_strFVSPrePostSeqNumRxPackageAssgnTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("FVS PRE-POST SEQNUM TREATMENT PACKAGE ASSIGNMENTS");
				
			
				if (this.m_strRxTable.Trim().Length == 0) 
				{
					
					MessageBox.Show("!!Could Not Locate Rx Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
				if (this.m_strFvsCatTable.Trim().Length == 0 && !ReferenceQueries.Scenario)
				{
					MessageBox.Show("!!Could Not Locate FVS Category Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
				if (this.m_strFvsSubCatTable.Trim().Length == 0 && !ReferenceQueries.Scenario)
				{
					MessageBox.Show("!!Could Not Locate FVS Subcategory Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
				if (this.m_strFvsCmdTable.Trim().Length == 0 && !ReferenceQueries.Scenario)
				{
					MessageBox.Show("!!Could Not Locate FVS Command Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
				if (this.m_strRxFvsCmdTable.Trim().Length == 0 && !ReferenceQueries.Scenario)
				{
					MessageBox.Show("!!Could Not Locate Rx FVS Command Assignments Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
				if (this.m_strRxHarvestCostColumnsTable.Trim().Length == 0 && ReferenceQueries._strScenarioType!="core")
				{
					MessageBox.Show("!!Could Not Locate Rx Harvest Cost Columns Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}

                if (this.m_strTreeSpcTable.Trim().Length == 0 && ReferenceQueries._strScenarioType != "core")
				{
					MessageBox.Show("!!Could Not Locate Tree Species Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}

				if (this.m_strFvsTreeSpcRefTable.Trim().Length == 0 && !ReferenceQueries.Scenario)
				{
					MessageBox.Show("!!Could Not Locate FVS Tree Species Reference Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
			}
			
			public string GetRxItemCategoryDescriptionSQL(string p_strRxId)
			{
				return "SELECT desc " + 
					   "FROM " + m_strFvsCatTable + " " + 
                       "WHERE MAX >= " + p_strRxId + " AND " + 
					   "MIN <= " + p_strRxId;
			}
			public string GetRxItemSubCategoryDescriptionSQL(string p_strRxId)
			{
				return "SELECT desc " + 
					   "FROM " + m_strFvsSubCatTable + " " + 
					   "WHERE MAX >= " + p_strRxId + " AND " + 
					   "MIN <= " + p_strRxId;
			}
			public string GetCategoryIdFromDescriptionSQL(string p_strDesc)
			{
				return "SELECT DISTINCT CSTR(catid) " + 
					   "FROM " + m_strFvsCatTable + " c " + 
					   "WHERE TRIM(UCASE(c.DESC)) = '" + p_strDesc.Trim().ToUpper() + "'";
			}
			public string GetSubCategoryIdFromDescriptionSQL(string p_strDesc)
			{
				return "SELECT DISTINCT CSTR(subcatid) " + 
					"FROM " + m_strFvsSubCatTable + " c " + 
					"WHERE TRIM(UCASE(c.DESC)) = '" + p_strDesc.Trim().ToUpper() + "'";
			}
			public string GetCategoryDescriptionFromCategoryIdSQL(string p_strCatId)
			{
				return "SELECT DISTINCT desc " + 
					"FROM " + m_strFvsCatTable + " c " + 
					"WHERE c.catid = " + p_strCatId;
			}
			public string GetSubCategoryDescriptionFromCategoryIdAndSubCategoryIdSQL(string p_strCatId,string p_strSubCatId)
			{
				return "SELECT DISTINCT desc " + 
					"FROM " + this.m_strFvsSubCatTable + " c " + 
					"WHERE c.catid = " + p_strCatId + " AND " + 
						  "c.subcatid = " + p_strSubCatId;
			}
			
			public string GetFVSCommandPropertiesSQL(string p_strCmd)
			{
				return "SELECT * " + 
					   "FROM " + m_strFvsCmdTable + " " + 
					   "WHERE TRIM(UCASE(fvscmd))='" + p_strCmd.Trim().ToUpper() + "'";
			}
			public string GetInsertRxSQL(string p_strFields, string p_strValues)
			{
				return "INSERT INTO " + this.m_strRxTable + " (" + p_strFields + ") VALUES (" + p_strValues + ")";
			}
			public string GetInsertRxFvsCmdSQL(string p_strFields, string p_strValues)
			{
				return "INSERT INTO " + this.m_strRxTable + " (" + p_strFields + ") VALUES (" + p_strValues + ")";
			}

			/// <summary>
			/// get the subcategories that are child members of the parent category.
			/// </summary>
			/// <param name="p_strCat"></param>
			/// <returns></returns>
			public string GetFVSSubCategoriesSQL(string p_strCat)
			{
				return "SELECT s.desc, s.min, s.max " + 
					   "FROM " + this.m_strFvsSubCatTable + " s," +
					             this.m_strFvsCatTable + " c " + 
					   "WHERE TRIM(UCASE(c.desc))='" + p_strCat.Trim().ToUpper() + "' AND " + 
					         "s.catid=c.catid";
			}
			/// <summary>
			/// return the query that will assign a variant to each rx package
			/// </summary>
			/// <param name="p_strPlotTable"></param>
			/// <param name="p_strRxPackageTable"></param>
			/// <returns></returns>
			static public string GetFVSVariantRxPackageSQL(string p_strPlotTable, string p_strRxPackageTable)
			{
				return "SELECT DISTINCT  a.fvs_variant,  b.rxpackage, b.rxcycle_length, b.simyear1_rx," + 
					                                                                   "b.simyear2_rx," + 
					                                                                   "b.simyear3_rx," + 
					                                                                   "b.simyear4_rx  " + 
					   "FROM " + p_strPlotTable + " a, " + 
					      "(SELECT rxpackage,simyear1_rx,simyear2_rx,simyear3_rx,simyear4_rx,rxcycle_length " + 
					       "FROM " + p_strRxPackageTable +  ") b " + 
					   "WHERE a.fvs_variant IS NOT NULL AND " + 
					          "LEN(TRIM(a.fvs_variant)) > 0 AND " + 
					          "b.rxpackage IS NOT NULL AND LEN(TRIM(b.rxpackage)) > 0;";
			}
			/// <summary>
			/// return the query that will assign a variant to each rx package that is in the list
			/// </summary>
			/// <param name="p_strPlotTable"></param>
			/// <param name="p_strRxPackageTable"></param>
			/// <param name="p_strRxPackageList"></param>
			/// <param name="p_strFilter">include a filter using alias (a) for plot table and alias (b) for rxpackage table</param>
			/// <returns></returns>
			static public string GetFVSVariantRxPackageSQL(string p_strPlotTable, string p_strRxPackageTable,string p_strFilter)
			{
				return "SELECT DISTINCT  a.fvs_variant,  b.rxpackage, b.rxcycle_length, b.simyear1_rx," + 
																					   "b.simyear2_rx," + 
																					   "b.simyear3_rx," + 
																					   "b.simyear4_rx  " + 
					"FROM " + p_strPlotTable + " a, " + 
						"(SELECT rxpackage,simyear1_rx,simyear2_rx,simyear3_rx,simyear4_rx,rxcycle_length " + 
						 "FROM " + p_strRxPackageTable +  ") b " + 
					"WHERE a.fvs_variant IS NOT NULL AND " + 
						 "LEN(TRIM(a.fvs_variant)) > 0 AND " + 
						"b.rxpackage IS NOT NULL AND LEN(TRIM(b.rxpackage)) > 0 AND " + 
					    p_strFilter;
			}
            /// <summary>
            ///  Assign a sequence number to each record of the FVS output table and group by standid,year
            /// </summary>
            /// <param name="p_strIntoTable"></param>
            /// <param name="p_strFVSOutputTable"></param>
            /// <param name="p_bAllColumns"></param>
            /// <returns></returns>
            static public string FVSOutputTable_PrePostGenericSQL(string p_strIntoTable, string p_strFVSOutputTable, bool p_bAllColumns)
            {
                string strSQL = "";
                if (p_bAllColumns)
                {
                    strSQL = "SELECT d.SeqNum,a.* " +
                              "FROM " + p_strFVSOutputTable + " a," +
                                 "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                          "b.standid, b.year " +
                                  "FROM " + p_strFVSOutputTable + " b," +
                                        "(SELECT standid,year " +
                                         "FROM " + p_strFVSOutputTable + ") c " +
                                 "WHERE b.standid=c.standid " +
                                 "GROUP BY b.standid,b.year) d " +
                              "WHERE a.standid=d.standid AND a.year=d.year";
                }
                else
                {
                    if (p_strIntoTable.Trim().Length > 0)
                    {
                        strSQL = "SELECT  SUM(IIF(a.year >= b.year,1,0)) AS SeqNum," +
                                          "a.standid, a.year " +
                                 "INTO " + p_strIntoTable + " " + 
                                 "FROM " + p_strFVSOutputTable + " a," +
                                      "(SELECT standid,year " +
                                       "FROM " + p_strFVSOutputTable + ") b " +
                                 "WHERE a.standid=b.standid " +
                                 "GROUP BY a.standid,a.year";
                    }
                    else
                    {
                        strSQL = "SELECT  SUM(IIF(a.year >= b.year,1,0)) AS SeqNum," +
                                          "a.standid, a.year " +
                                 "FROM " + p_strFVSOutputTable + " a," +
                                      "(SELECT standid,year " +
                                       "FROM " + p_strFVSOutputTable + ") b " +
                                 "WHERE a.standid=b.standid " +
                                 "GROUP BY a.standid,a.year";
                    }
                }
                return strSQL;
               
            }
            static public string[] FVSOutputTable_PrePostPotFireBaseYearSQL(string p_strPotFireBaseYearTable, string p_strPotFireTable,string p_strWorkTableName)
            {
                string[] strSQL = new string[9];

                //create a baseyear table that contains only standid's that are in the standard table
                strSQL[0] = "SELECT DISTINCT a.* " +
                            "INTO tempBASEYEAR " +
                            "FROM " + p_strPotFireBaseYearTable + " a " +
                            "INNER JOIN " + p_strPotFireTable + " b ON a.standid=b.standid";

                //get the potfire base year records into baseyear temp table
                strSQL[1] = "SELECT 'Y' AS BASEYEAR_YN,a.* " +
                          "INTO BASEYEAR " +
                          "FROM tempBASEYEAR a," +
                            "(SELECT STANDID, MIN([YEAR]) AS BASEYEAR, 'Y' AS BASEYEAR_YN " +
                             "FROM tempBASEYEAR " +
                            "GROUP BY STANDID) b " +
                         "WHERE a.standid = b.standid AND a.year=b.baseyear";

                //get fvs_potfire  records into nonbaseyear temp table and increment the year by 1
                strSQL[2] = "SELECT 'N' AS BASEYEAR_YN,(a.[YEAR] + 1) AS [NEWYEAR], a.* " +
                            "INTO NONBASEYEAR " +
                            "FROM " + p_strPotFireTable + " a ";

                //update the year column to the newyear from the previous step
                strSQL[3] = "UPDATE NONBASEYEAR SET [YEAR]=NEWYEAR";

                //drop the newyear column
                strSQL[4] = "ALTER TABLE NONBASEYEAR DROP COLUMN NEWYEAR";

                strSQL[5] = "SELECT * INTO " + p_strWorkTableName + " FROM BASEYEAR";

                strSQL[6] = "INSERT INTO " + p_strWorkTableName + " SELECT * FROM NONBASEYEAR";

                strSQL[7] = "DROP TABLE NONBASEYEAR";

                strSQL[8] = "DROP TABLE BASEYEAR";

                strSQL[8] = "DROP TABLE tempBASEYEAR";

                return strSQL;
            }
            static public string[] FVSOutputTable_PrePostPotFireBaseYearIDColumnSQL(string p_strPotFireBaseYearTable, string p_strPotFireTable, string p_strWorkTableName)
            {
                string[] strSQL = new string[12];

                //create a baseyear table that contains only standid's that are in the standard table
                strSQL[0] = "SELECT DISTINCT a.* " +
                            "INTO tempBASEYEAR " +
                            "FROM " + p_strPotFireBaseYearTable + " a " +
                            "INNER JOIN " + p_strPotFireTable + " b ON a.standid=b.standid";

                //get the potfire base year records into baseyear temp table
                strSQL[1] = "SELECT 'Y' AS BASEYEAR_YN,a.* " +
                          "INTO BASEYEAR " +
                          "FROM tempBASEYEAR a," +
                            "(SELECT STANDID, MIN([YEAR]) AS BASEYEAR, 'Y' AS BASEYEAR_YN " +
                             "FROM tempBASEYEAR " +
                            "GROUP BY STANDID) b " +
                         "WHERE a.standid = b.standid AND a.year=b.baseyear";

                //get fvs_potfire  records into nonbaseyear temp table and increment the year by 1
                strSQL[2] = "SELECT 'N' AS BASEYEAR_YN,(a.[YEAR] + 1) AS [NEWYEAR], a.* " +
                            "INTO NONBASEYEAR " +
                            "FROM " + p_strPotFireTable + " a ";

                //update the year column to the newyear from the previous step
                strSQL[3] = "UPDATE NONBASEYEAR SET [YEAR]=NEWYEAR";

                //drop the newyear column
                strSQL[4] = "ALTER TABLE NONBASEYEAR DROP COLUMN NEWYEAR";



                //get the maximum id value from the nonbaseyear table, create a rownumber for each row in the baseyear table,
                //and add the rownumber to the maxid to get a unique id
                strSQL[5] = "SELECT a.*,b.rownumber + c.maxid AS tempid INTO " + p_strWorkTableName + " FROM BASEYEAR a,";
                strSQL[5] = strSQL[5] +
                            "(" + Queries.Utilities.AssignRowNumberToEachRow("", "BASEYEAR", "STANDID", "ROWNUMBER") +
                            ") b," +
                            "(SELECT MAX(ID) AS maxid FROM NONBASEYEAR) c " +
                            "WHERE a.standid=b.standid";

                //update in the baseyear temp table with the new id
                strSQL[6] = "UPDATE " + p_strWorkTableName + " SET ID=TEMPID";


                strSQL[7] = "ALTER TABLE " + p_strWorkTableName + " DROP COLUMN TEMPID";

                strSQL[8] = "INSERT INTO " + p_strWorkTableName + " SELECT * FROM NONBASEYEAR";

              
                strSQL[9] = "DROP TABLE NONBASEYEAR";

                strSQL[10] = "DROP TABLE BASEYEAR";

                strSQL[11] = "DROP TABLE tempBASEYEAR";

                return strSQL;
            }
            /// <summary>
            /// Assign a sequence number to each record of the FVS output table (FVS_STRCLASS) and group by standid,year and removal code.
            /// </summary>
            /// <param name="p_strIntoTable"></param>
            /// <param name="p_strFVSOutputTable">FVS_STRCLASS table name</param>
            /// <param name="p_bAllColumns"></param>
            /// <param name="p_strRemovalCode">Values: 0 (before removal) or 1 (after removal)</param>
            /// <returns></returns>
            static public string FVSOutputTable_PrePostStrClassSQL(string p_strIntoTable,string p_strFVSOutputTable,bool p_bAllColumns,string p_strRemovalCode)
            {
                string strSQL = "";
                if (p_bAllColumns)
                {
                    strSQL = "SELECT d.seqnum, z.* FROM " + p_strFVSOutputTable + " z," +
                                    "(SELECT  SUM(IIF(a.year >= b.year AND " + 
                                                     "a.removal_code=" + p_strRemovalCode + ",1,0)) AS SeqNum," + 
                                              "a.standid, a.year,b.removal_code " + 
                                     "FROM fvs_strclass a," +
                                        "(SELECT standid,year,removal_code " + 
                                         "FROM " + p_strFVSOutputTable + " " + 
                                         "WHERE removal_code=" + p_strRemovalCode + ") b " +
                                    "WHERE a.standid=b.standid AND " + 
                                          "a.removal_code=b.removal_code AND " + 
                                          "b.removal_code=" + p_strRemovalCode + " " +
                                    "GROUP BY a.standid,a.year,b.removal_code) d " +
                             "WHERE z.standid=d.standid AND z.year=d.year AND z.removal_code=d.removal_code";
                }
                else
                {
                    if (p_strIntoTable.Trim().Length > 0)
                    {
                        strSQL = "SELECT  SUM(IIF(a.year >= b.year,1,0)) AS SeqNum," +
                                          "a.standid, a.year " +
                                 "INTO " + p_strIntoTable + " " +
                                 "FROM " + p_strFVSOutputTable + " a," +
                                      "(SELECT standid,year " +
                                       "FROM " + p_strFVSOutputTable + ") b " +
                                 "WHERE a.standid=b.standid " +
                                 "GROUP BY a.standid,a.year";
                    }
                    else
                    {
                        strSQL = "SELECT  SUM(IIF(a.year >= b.year,1,0)) AS SeqNum," +
                                          "a.standid, a.year " +
                                 "FROM " + p_strFVSOutputTable + " a," +
                                      "(SELECT standid,year " +
                                       "FROM " + p_strFVSOutputTable + ") b " +
                                 "WHERE a.standid=b.standid " +
                                 "GROUP BY a.standid,a.year";
                    }
                }
                return strSQL;

            }
            /// <summary>
            /// SQL for creating the sequence number configuration table from the FVS Output table.
            /// </summary>
            /// <param name="p_strIntoTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <param name="p_strFVSOutputTable"></param>
            /// <param name="p_bAllColumns"></param>
            /// <returns></returns>
            static public string FVSOutputTable_AuditPrePostGenericSQL(string p_strIntoTable, string p_strFVSOutputTable, bool p_bAllColumns)
            {
                string strSQL = "";
               
                    if (p_strIntoTable.Trim().Length > 0)
                    {
                        strSQL = "SELECT d.SeqNum,a.standid,a.year," + 
                                       "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," + 
                                       "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," + 
                                       "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                       "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN " + 
                                 "INTO " + p_strIntoTable + " " + 
                                 "FROM " + p_strFVSOutputTable + " a," +
                                     "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                              "b.standid, b.year " +
                                      "FROM " + p_strFVSOutputTable + " b," +
                                            "(SELECT standid,year " +
                                             "FROM " + p_strFVSOutputTable + ") c " +
                                     "WHERE b.standid=c.standid " +
                                     "GROUP BY b.standid,b.year) d " +
                                  "WHERE a.standid=d.standid AND a.year=d.year";
                    }
                    else
                    {
                        strSQL = "SELECT d.SeqNum,a.standid,a.year," +
                                       "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," +
                                       "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," +
                                       "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                       "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN " + 
                                 "FROM " + p_strFVSOutputTable + " a," +
                                    "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                             "b.standid, b.year " +
                                     "FROM " + p_strFVSOutputTable + " b," +
                                           "(SELECT standid,year " +
                                            "FROM " + p_strFVSOutputTable + ") c " +
                                    "WHERE b.standid=c.standid " +
                                    "GROUP BY b.standid,b.year) d " +
                                 "WHERE a.standid=d.standid AND a.year=d.year " + 
                                 "ORDER BY a.standid,d.SeqNum";
                    }
               
                    return strSQL;

            }
            /// <summary>
            /// Update SQL for assigning the sequence number to the PRE or POST cycle.
            /// </summary>
            /// <param name="p_oItem"></param>
            /// <param name="p_strUpdateTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <returns></returns>
            static public string FVSOutputTable_AuditUpdatePrePostGenericSQL(FVSPrePostSeqNumItem p_oItem,string p_strUpdateTable)
            {
                string strSQL="";
                 if (p_oItem.RxCycle1PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle1PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE1_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle1PreSeqNum + ",'Y','N'),";

                }
                 if (p_oItem.RxCycle2PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle2PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE2_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle2PreSeqNum + ",'Y','N'),";
                }

                 if (p_oItem.RxCycle3PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle3PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE3_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle3PreSeqNum + ",'Y','N'),";
                }
                 if (p_oItem.RxCycle4PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle4PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE4_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle4PreSeqNum + ",'Y','N'),";
                }
                 if (p_oItem.RxCycle1PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle1PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE1_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle1PostSeqNum + ",'Y','N'),";

                }
                 if (p_oItem.RxCycle2PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle2PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE2_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle2PostSeqNum + ",'Y','N'),";
                }

                 if (p_oItem.RxCycle3PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle3PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE3_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle3PostSeqNum + ",'Y','N'),";
                }
                 if (p_oItem.RxCycle4PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle4PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE4_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle4PostSeqNum + ",'Y','N'),";
                }
                if (strSQL.Trim().Length > 0)
                {
                    strSQL = strSQL.Substring(0, strSQL.Length - 1);
                    strSQL = "UPDATE " + p_strUpdateTable + " " +
                                      "SET " + strSQL;
                }
                return strSQL;
            }
            /// <summary>
            /// Update SQL for assigning the sequence number to the PRE or POST cycle.
            /// </summary>
            /// <param name="p_oItem"></param>
            /// <param name="p_strUpdateTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <returns></returns>
            static public string FVSOutputTable_AuditUpdatePrePostStrClassSQL(FVSPrePostSeqNumItem p_oItem, string p_strUpdateTable)
            {
                string strSQL = "";
                if (p_oItem.RxCycle1PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle1PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    if (p_oItem.RxCycle1PreStrClassBeforeTreeRemovalYN == "N")
                    {
                        strSQL = strSQL + "CYCLE1_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle1PreSeqNum + " AND removal_code=1 ,'Y','N'),";
                    }
                    else
                    {
                        strSQL = strSQL + "CYCLE1_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle1PreSeqNum + " AND removal_code=0 ,'Y','N'),";
                    }

                }
                if (p_oItem.RxCycle2PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle2PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    if (p_oItem.RxCycle2PreStrClassBeforeTreeRemovalYN == "N")
                    {
                        strSQL = strSQL + "CYCLE2_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle2PreSeqNum + " AND removal_code=1 ,'Y','N'),";
                    }
                    else
                    {
                        strSQL = strSQL + "CYCLE2_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle2PreSeqNum + " AND removal_code=0 ,'Y','N'),";
                    }
                }

                if (p_oItem.RxCycle3PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle3PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    if (p_oItem.RxCycle3PreStrClassBeforeTreeRemovalYN == "N")
                    {
                        strSQL = strSQL + "CYCLE3_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle3PreSeqNum + " AND removal_code=1 ,'Y','N'),";
                    }
                    else
                    {
                        strSQL = strSQL + "CYCLE3_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle3PreSeqNum + " AND removal_code=0 ,'Y','N'),";
                    }
                }
                if (p_oItem.RxCycle4PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle4PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    if (p_oItem.RxCycle4PreStrClassBeforeTreeRemovalYN == "N")
                    {
                        strSQL = strSQL + "CYCLE4_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle4PreSeqNum + " AND removal_code=1 ,'Y','N'),";
                    }
                    else
                    {
                        strSQL = strSQL + "CYCLE4_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle4PreSeqNum + " AND removal_code=0 ,'Y','N'),";
                    }
                }
                if (p_oItem.RxCycle1PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle1PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    if (p_oItem.RxCycle1PostStrClassBeforeTreeRemovalYN == "N")
                    {
                        strSQL = strSQL + "CYCLE1_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle1PostSeqNum + " AND removal_code=1 ,'Y','N'),";
                    }
                    else
                    {
                        strSQL = strSQL + "CYCLE1_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle1PostSeqNum + " AND removal_code=0 ,'Y','N'),";
                    }

                }
                if (p_oItem.RxCycle2PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle2PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    if (p_oItem.RxCycle2PostStrClassBeforeTreeRemovalYN == "N")
                    {
                        strSQL = strSQL + "CYCLE2_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle2PostSeqNum + " AND removal_code=1 ,'Y','N'),";
                    }
                    else
                    {
                        strSQL = strSQL + "CYCLE2_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle2PostSeqNum + " AND removal_code=0 ,'Y','N'),";
                    }
                }

                if (p_oItem.RxCycle3PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle3PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    if (p_oItem.RxCycle3PostStrClassBeforeTreeRemovalYN == "N")
                    {
                        strSQL = strSQL + "CYCLE3_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle3PostSeqNum + " AND removal_code=1 ,'Y','N'),";
                    }
                    else
                    {
                        strSQL = strSQL + "CYCLE3_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle3PostSeqNum + " AND removal_code=0 ,'Y','N'),";
                    }
                }
                if (p_oItem.RxCycle4PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle4PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    if (p_oItem.RxCycle4PostStrClassBeforeTreeRemovalYN == "N")
                    {
                        strSQL = strSQL + "CYCLE4_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle4PostSeqNum + " AND removal_code=1 ,'Y','N'),";
                    }
                    else
                    {
                        strSQL = strSQL + "CYCLE4_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle4PostSeqNum + " AND removal_code=0 ,'Y','N'),";
                    }
                }
                if (strSQL.Trim().Length > 0)
                {
                    strSQL = strSQL.Substring(0, strSQL.Length - 1);
                    strSQL = "UPDATE " + p_strUpdateTable + " " +
                                      "SET " + strSQL;
                }
                return strSQL;
            }
            /// <summary>
            /// SQL for creating the sequence number configuration table from the FVS Output table FVS_STRCLASS.
            /// </summary>
            /// <param name="p_strIntoTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <param name="p_strFVSOutputTable"></param>
            /// <param name="p_bAllColumns"></param>
            /// <returns></returns>
            static public string FVSOutputTable_AuditPrePostStrClassSQL(string p_strIntoTable, string p_strFVSOutputTable, bool p_bAllColumns)
            {
                string strSQL = "";

                if (p_strIntoTable.Trim().Length > 0)
                {
                    strSQL = "SELECT d.SeqNum,a.standid,a.year,a.removal_code," +
                                   "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," +
                                   "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," +
                                   "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                   "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN " +
                             "INTO " + p_strIntoTable + " " +
                             "FROM " + p_strFVSOutputTable + " a," +
                                 "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                          "b.standid, b.year,b.removal_code " +
                                  "FROM " + p_strFVSOutputTable + " b," +
                                        "(SELECT standid,year,removal_code " +
                                         "FROM " + p_strFVSOutputTable + ") c " +
                                 "WHERE b.standid=c.standid AND b.removal_code=c.removal_code " +
                                 "GROUP BY b.standid,b.year,b.removal_code) d " +
                              "WHERE a.standid=d.standid AND a.year=d.year AND a.removal_code=d.removal_code";
                }
                else
                {
                    strSQL = "SELECT d.SeqNum,a.standid,a.year,a.removal_code," +
                                   "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," +
                                   "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," +
                                   "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                   "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN " +
                             "FROM " + p_strFVSOutputTable + " a," +
                                 "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                          "b.standid, b.year,b.removal_code " +
                                  "FROM " + p_strFVSOutputTable + " b," +
                                        "(SELECT standid,year,removal_code " +
                                         "FROM " + p_strFVSOutputTable + ") c " +
                                 "WHERE b.standid=c.standid AND b.removal_code=c.removal_code " +
                                 "GROUP BY b.standid,b.year,b.removal_code) d " +
                             "WHERE a.standid=d.standid AND a.year=d.year AND a.removal_code=d.removal_code";
                }

                return strSQL;

            }
            /// <summary>
            /// SQL for creating the sequence number configuration table from the FVS Output table FVS_SUMMARY.
            /// </summary>
            /// <param name="p_strIntoTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <param name="p_strFVSOutputTable"></param>
            /// <param name="p_bAllColumns"></param>
            /// <returns></returns>
            static public string[] FVSOutputTable_AuditPrePostFvsStrClassUsingFVSSummarySQL(string p_strIntoTable, string p_strFVSSummaryTable, bool p_bAllColumns)
            {
                string[] strSQL = new string[2];

                if (p_strIntoTable.Trim().Length > 0)
                {
                 strSQL[0] = "SELECT d.SeqNum,a.standid,a.year,0 AS removal_code," +
                                   "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," +
                                   "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," +
                                   "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                   "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN " +
                             "INTO " + p_strIntoTable + " " +
                             "FROM " + p_strFVSSummaryTable + " a," +
                                 "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                          "b.standid, b.year " +
                                  "FROM " + p_strFVSSummaryTable + " b," +
                                        "(SELECT standid,year " +
                                         "FROM " + p_strFVSSummaryTable + ") c " +
                                 "WHERE b.standid=c.standid " +
                                 "GROUP BY b.standid,b.year) d " +
                              "WHERE a.standid=d.standid AND a.year=d.year";

                 strSQL[1] = "INSERT INTO " + p_strIntoTable + " " +
                             "SELECT d.SeqNum,a.standid,a.year,1 AS removal_code," +
                                   "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," +
                                   "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," +
                                   "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                   "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN " +
                             "FROM " + p_strFVSSummaryTable + " a," +
                                 "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                          "b.standid, b.year " +
                                  "FROM " + p_strFVSSummaryTable + " b," +
                                        "(SELECT standid,year " +
                                         "FROM " + p_strFVSSummaryTable + ") c " +
                                 "WHERE b.standid=c.standid " +
                                 "GROUP BY b.standid,b.year) d " +
                             "WHERE a.standid=d.standid AND a.year=d.year";
                }
                else
                {
                    strSQL[0] = "SELECT d.SeqNum,a.standid,a.year,0 AS removal_code," +
                                   "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," +
                                   "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," +
                                   "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                   "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN " +
                             "FROM " + p_strFVSSummaryTable + " a," +
                                 "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                          "b.standid, b.year " +
                                  "FROM " + p_strFVSSummaryTable + " b," +
                                        "(SELECT standid,year " +
                                         "FROM " + p_strFVSSummaryTable + ") c " +
                                 "WHERE b.standid=c.standid " +
                                 "GROUP BY b.standid,b.year) d " +
                             "WHERE a.standid=d.standid AND a.year=d.year";

                    strSQL[1] = "SELECT d.SeqNum,a.standid,a.year,1 AS removal_code," +
                                   "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," +
                                   "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," +
                                   "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                   "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN " +
                             "FROM " + p_strFVSSummaryTable + " a," +
                                 "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                          "b.standid, b.year " +
                                  "FROM " + p_strFVSSummaryTable + " b," +
                                        "(SELECT standid,year " +
                                         "FROM " + p_strFVSSummaryTable + ") c " +
                                 "WHERE b.standid=c.standid " +
                                 "GROUP BY b.standid,b.year) d " +
                             "WHERE a.standid=d.standid AND a.year=d.year";
                }

                return strSQL;

            }

            /// <summary>
            /// Audit to identify assigned sequence numbers that cannot be found in the Sequence Number Matrix table (FVSTableName_PREPOST_SEQNUM_MATRIX)
            /// </summary>
            /// <param name="p_oFVSPrePostSeqNumItem"></param>
            /// <param name="p_strIntoTable"></param>
            /// <param name="p_strSourceTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <returns></returns>
            static public string FVSOutputTable_AuditSelectIntoPrePostSeqNumCount(FVSPrePostSeqNumItem p_oFVSPrePostSeqNumItem,string p_strIntoTable,string p_strSourceTable)
            {
                string strSQL = "";
                int x;
                int z = 0;
               
                string strAlpha = "cdefghij";
                int intAlias = 0;
                string strSelectColumns = "a.standid,b.totalrows,";
                
                //cycle 1 seqnum
                if (p_oFVSPrePostSeqNumItem.RxCycle1PreSeqNum.Trim().Length > 0 && p_oFVSPrePostSeqNumItem.RxCycle1PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + " (SELECT " +
                                         "SUM(IIF(SeqNum=" + p_oFVSPrePostSeqNumItem.RxCycle1PreSeqNum + ",1,0)) AS pre_cycle1rows," +
                                         "STANDID " +
                                       "FROM " + p_strSourceTable + " " +
                                       "GROUP BY STANDID) " + strAlpha.Substring(intAlias, 1) + ",";
                    strSelectColumns = strSelectColumns + strAlpha.Substring(intAlias, 1) + ".pre_cycle1rows,";
                    intAlias++;
                }
                if (p_oFVSPrePostSeqNumItem.RxCycle1PostSeqNum.Trim().Length > 0 && p_oFVSPrePostSeqNumItem.RxCycle1PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + " (SELECT " +
                                         "SUM(IIF(SeqNum=" + p_oFVSPrePostSeqNumItem.RxCycle1PostSeqNum + ",1,0)) AS post_cycle1rows," +
                                         "STANDID " +
                                       "FROM " + p_strSourceTable + " " +
                                       "GROUP BY STANDID) " + strAlpha.Substring(intAlias, 1) + ",";
                    strSelectColumns = strSelectColumns + strAlpha.Substring(intAlias, 1) + ".post_cycle1rows,";
                    intAlias++;
                }
                //cycle 2 seqnum
                if (p_oFVSPrePostSeqNumItem.RxCycle2PreSeqNum.Trim().Length > 0 && p_oFVSPrePostSeqNumItem.RxCycle2PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + " (SELECT " +
                                         "SUM(IIF(SeqNum=" + p_oFVSPrePostSeqNumItem.RxCycle2PreSeqNum + ",1,0)) AS pre_cycle2rows," +
                                         "STANDID " +
                                       "FROM " + p_strSourceTable + " " +
                                       "GROUP BY STANDID) " + strAlpha.Substring(intAlias, 1) + ",";
                    strSelectColumns = strSelectColumns + strAlpha.Substring(intAlias, 1) + ".pre_cycle2rows,";
                    intAlias++;
                }
                if (p_oFVSPrePostSeqNumItem.RxCycle2PostSeqNum.Trim().Length > 0 && p_oFVSPrePostSeqNumItem.RxCycle2PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + " (SELECT " +
                                         "SUM(IIF(SeqNum=" + p_oFVSPrePostSeqNumItem.RxCycle2PostSeqNum + ",1,0)) AS post_cycle2rows," +
                                         "STANDID " +
                                       "FROM " + p_strSourceTable + " " +
                                       "GROUP BY STANDID) " + strAlpha.Substring(intAlias, 1) + ",";
                    strSelectColumns = strSelectColumns + strAlpha.Substring(intAlias, 1) + ".post_cycle2rows,";
                    intAlias++;
                }
                //cycle 3 seqnum
                if (p_oFVSPrePostSeqNumItem.RxCycle3PreSeqNum.Trim().Length > 0 && p_oFVSPrePostSeqNumItem.RxCycle3PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + " (SELECT " +
                                         "SUM(IIF(SeqNum=" + p_oFVSPrePostSeqNumItem.RxCycle3PreSeqNum + ",1,0)) AS pre_cycle3rows," +
                                         "STANDID " +
                                       "FROM " + p_strSourceTable + " " +
                                       "GROUP BY STANDID) " + strAlpha.Substring(intAlias, 1) + ",";
                    strSelectColumns = strSelectColumns + strAlpha.Substring(intAlias, 1) + ".pre_cycle3rows,";
                    intAlias++;
                }
                if (p_oFVSPrePostSeqNumItem.RxCycle3PostSeqNum.Trim().Length > 0 && p_oFVSPrePostSeqNumItem.RxCycle3PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + " (SELECT " +
                                         "SUM(IIF(SeqNum=" + p_oFVSPrePostSeqNumItem.RxCycle3PostSeqNum + ",1,0)) AS post_cycle3rows," +
                                         "STANDID " +
                                       "FROM " + p_strSourceTable + " " +
                                       "GROUP BY STANDID) " + strAlpha.Substring(intAlias, 1) + ",";
                    strSelectColumns = strSelectColumns + strAlpha.Substring(intAlias, 1) + ".post_cycle3rows,";
                    intAlias++;
                }
                //cycle 4 seqnum
                if (p_oFVSPrePostSeqNumItem.RxCycle4PreSeqNum.Trim().Length > 0 && p_oFVSPrePostSeqNumItem.RxCycle4PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + " (SELECT " +
                                         "SUM(IIF(SeqNum=" + p_oFVSPrePostSeqNumItem.RxCycle4PreSeqNum + ",1,0)) AS pre_cycle4rows," +
                                         "STANDID " +
                                       "FROM " + p_strSourceTable + " " +
                                       "GROUP BY STANDID) " + strAlpha.Substring(intAlias, 1) + ",";
                    strSelectColumns = strSelectColumns + strAlpha.Substring(intAlias, 1) + ".pre_cycle4rows,";
                    intAlias++;
                }
                if (p_oFVSPrePostSeqNumItem.RxCycle4PostSeqNum.Trim().Length > 0 && p_oFVSPrePostSeqNumItem.RxCycle4PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + " (SELECT " +
                                         "SUM(IIF(SeqNum=" + p_oFVSPrePostSeqNumItem.RxCycle4PostSeqNum + ",1,0)) AS post_cycle4rows," +
                                         "STANDID " +
                                       "FROM " + p_strSourceTable + " " +
                                       "GROUP BY STANDID) " + strAlpha.Substring(intAlias, 1) + ",";
                    strSelectColumns = strSelectColumns + strAlpha.Substring(intAlias, 1) + ".post_cycle4rows,";
                    intAlias++;
                }
                strSQL = strSQL.Substring(0, strSQL.Length - 1);
                strSelectColumns = strSelectColumns.Substring(0, strSelectColumns.Length - 1);

                strSQL = "SELECT DISTINCT " + strSelectColumns + " " +
                                 "INTO " + p_strIntoTable + " " +
                                 "FROM " + p_strSourceTable + " a," +
                                    "(SELECT COUNT(*) AS totalrows," +
                                            "STANDID " +
                                     "FROM " + p_strSourceTable + " " +
                                 "GROUP BY standid) b," +
                                 strSQL + " " +
                                 "WHERE a.standid=b.standid AND ";

                for (x = 0; x <= intAlias - 1; x++)
                {
                    strSQL = strSQL + "a.standid=" + strAlpha.Substring(x, 1) + ".standid AND ";
                }

                strSQL = strSQL.Substring(0, strSQL.Length - 5);
                return strSQL;
            }
            /// <summary>
            /// Load data that will show which fvs_summary SeqNum are associated with harvested tree records
            /// </summary>
            /// <param name="p_strIntoTable1">Table to display to user</param>
            /// <param name="p_strIntoTable2">Temporary table</param>
            /// <param name="p_strFVSOutputSummaryTable">FVS output summary table</param>
            /// <param name="p_strFVSOutputTreeTable">FVS output tree cut table</param>
            /// <returns></returns>
            static public string[] FVSOutputTable_TreeHarvestSeqNumSQL(string p_strIntoTable1,string p_strIntoTable2, string p_strFVSOutputSummaryTable, string p_strFVSOutputTreeTable)
            {
                string[] strSQL = new string[4];

                strSQL[0] = "SELECT  SUM(IIF(a.year >= b.year,1,0)) AS FVS_Summary_SeqNum," +
                                          "a.standid, a.year,0 AS treecount  " +
                                 "INTO " + p_strIntoTable1 + " " + 
                                 "FROM " + p_strFVSOutputSummaryTable + " a," +
                                      "(SELECT standid,year " +
                                       "FROM " + p_strFVSOutputSummaryTable + ") b " +
                                 "WHERE a.standid=b.standid " +
                                 "GROUP BY a.standid,a.year";

                strSQL[1] = "SELECT SUM(1) AS TREECOUNT, STANDID, YEAR INTO " + p_strIntoTable2 + " " +
                            "FROM " + p_strFVSOutputTreeTable + " " + 
                            "GROUP BY STANDID,YEAR";

                strSQL[2] = "UPDATE " + p_strIntoTable1 + " a " +
                            "INNER JOIN " + p_strIntoTable2 + " b " +
                            "ON a.standid=b.standid AND a.year=b.year " +
                            "SET a.treecount=b.treecount";

                strSQL[3] = "DROP TABLE " + p_strIntoTable2;

                return strSQL;

            }
            /// <summary>
            /// Load data that will show which fvs_summary SeqNum are associated with target table records
            /// </summary>
            /// <param name="p_strIntoTable1">Table to display to user</param>
            /// <param name="p_strIntoTable2">Temporary table</param>
            /// <param name="p_strFVSOutputSummaryTable">FVS output summary table</param>
            /// <param name="p_strFVSOutputTreeTable">FVS output table</param>
            /// <returns></returns>
            static public string[] FVSOutputTable_AssignSummarySeqNumSQL(string p_strIntoTable1,string p_strIntoTable2, string p_strFVSOutputSummaryTable, string p_strFVSOutputTable)
            {
                string[] strSQL = new string[4];

                strSQL[0] = "SELECT  SUM(IIF(a.year >= b.year,1,0)) AS FVS_Summary_SeqNum," +
                                          "a.standid, a.year,0 AS rowcount  " +
                                 "INTO " + p_strIntoTable1 + " " + 
                                 "FROM " + p_strFVSOutputSummaryTable + " a," +
                                      "(SELECT standid,year " +
                                       "FROM " + p_strFVSOutputSummaryTable + ") b " +
                                 "WHERE a.standid=b.standid " +
                                 "GROUP BY a.standid,a.year";

                strSQL[1] = "SELECT SUM(1) AS ROWCOUNT, STANDID, YEAR INTO " + p_strIntoTable2 + " " +
                            "FROM " + p_strFVSOutputTable + " " + 
                            "GROUP BY STANDID,YEAR";

                strSQL[2] = "UPDATE " + p_strIntoTable1 + " a " +
                            "INNER JOIN " + p_strIntoTable2 + " b " +
                            "ON a.standid=b.standid AND a.year=b.year " +
                            "SET a.rowcount=b.rowcount";

                strSQL[3] = "DROP TABLE " + p_strIntoTable2;

                return strSQL;

            }

            static public string[] FVSOutputTable_AuditFVSSummaryTableRowCountsSQL(string p_strIntoTable, string p_strSummaryAuditTable, string p_strFVSOutputTable, string p_strRowCountColumn)
            {
                string[] strSQL = new string[4];

                strSQL[0] =  "SELECT DISTINCT a.standid,a.year,b.rowcount " + 
                             "INTO " + p_strIntoTable + " " + 
                             "FROM " + p_strFVSOutputTable + " a," +
                               "(SELECT COUNT(*) as rowcount,standid,year " + 
                                "FROM " + p_strFVSOutputTable + " " + 
                                "GROUP BY standid,year) b " + 
                             "WHERE a.standid=b.standid AND a.year=b.year";

               strSQL[1] = "UPDATE " + p_strSummaryAuditTable + " a " +  
                           "INNER JOIN " + p_strIntoTable + " b " + 
                           "ON a.standid=b.standid AND " + 
	                          "a.year=b.year " + 
                           "SET a." + p_strRowCountColumn + "=b.rowcount";

              strSQL[2] = "UPDATE " + p_strSummaryAuditTable + " " + 
                          "SET " + p_strRowCountColumn + "=0 " + 
                          "WHERE " + p_strRowCountColumn + " IS NULL";

              strSQL[3] = "DROP TABLE " + p_strIntoTable;

              return strSQL;
            }
			static public string FVSOutputTable_AuditSelectIntoCyleYearSQL(string p_strIntoTable, string p_strFVSOutputTable, int p_intCycleLength)
			{
					return "SELECT DISTINCT a.standid, b.year1, 'Y' AS year1_exists_yn," + 
											"b.year1 + " + p_intCycleLength.ToString() +  " AS year2," + 
											"'N' AS year2_exists_yn," + 
											"b.year1 + (" + p_intCycleLength.ToString() +  " * 2) AS year3," + 
											"'N' AS year3_exists_yn," + 
											"b.year1 + (" + p_intCycleLength.ToString() +  " * 3) AS year4," + 
											"'N' AS year4_exists_yn," + 
											"b.year1 + (" + p_intCycleLength.ToString() +  " * 4) AS year5," + 
											"'N' AS year5_exists_yn" + " " + 
						   "INTO " + p_strIntoTable + " " + 
						   "FROM " + p_strFVSOutputTable + " a," + 
								"(SELECT MIN([YEAR]) AS year1, standid " + 
						         "FROM " + p_strFVSOutputTable + " " + 
						         "GROUP BY standid) b" + " " + 
						   "WHERE a.standid = b.standid AND a.[year]=b.year1";
			}
			static public string FVSOutputTable_AuditUpdateCyleYearSQL(string p_strUpdateTableName, string p_strFVSOutputTable,string p_strCycle)
			{
				return "UPDATE " + p_strUpdateTableName + " a " + 
					   "INNER JOIN " + p_strFVSOutputTable + " b " + 
					   "ON a.standid=b.standid AND a.year" + p_strCycle.Trim() + "= b.year " + 
					   "SET year" + p_strCycle.Trim() + "_exists_yn=IIF(b.year=a.year" + p_strCycle.Trim() + ",'Y','N')";
			}
			static public string FVSOutputTable_AuditInsertUpdateCycleYearSQL(string p_strInsertTable, string p_strSourceTable,string p_strCycle)
			{
				return "INSERT INTO " + p_strInsertTable + " " + 
					      "(standid,`year`, cycle) " + 
					      "SELECT standid,year" + p_strCycle + " AS `year`," + 
								p_strCycle + " AS cycle " + 
					      "FROM " + p_strSourceTable;
			}
			static public string FVSOutputTable_AuditSelectIntoRxPostCyleYearSQL(string p_strIntoTable, 
															   string p_strFVSOutputTable,
				                                               string p_strFVSOutputCycleYearTable, 
				                                               string p_strCycle)
			{

				switch (p_strCycle)
				{
				    case "1":
						return  "SELECT a.standid, MIN(a.year) AS post_year" + p_strCycle + " " + 
							"INTO "  + p_strIntoTable + " " + 
							"FROM  " + p_strFVSOutputTable + " a, " + 
							p_strFVSOutputCycleYearTable + " b " + 
							"WHERE a.standid=b.standid AND " + 
							"a.year > b.year1 GROUP BY a.standid";
					case "2":
						return  "SELECT a.standid, MIN(a.year) AS post_year" + p_strCycle + " " + 
							"INTO "  + p_strIntoTable + " " + 
							"FROM  " + p_strFVSOutputTable + " a, " + 
							p_strFVSOutputCycleYearTable + " b " + 
							"WHERE a.standid=b.standid AND " + 
							"a.year <> b.year1 AND a.year > b.year2 GROUP BY a.standid";
					case "3":
						return  "SELECT a.standid, MIN(a.year) AS post_year" + p_strCycle + " " + 
							"INTO "  + p_strIntoTable + " " + 
							"FROM  " + p_strFVSOutputTable + " a, " + 
							p_strFVSOutputCycleYearTable + " b " + 
							"WHERE a.standid=b.standid AND " + 
						    "a.year <> b.year1 AND a.year <> b.year2 AND a.year > b.year3 GROUP BY a.standid";
					case "4":
						return  "SELECT a.standid, MIN(a.year) AS post_year" + p_strCycle + " " + 
							"INTO "  + p_strIntoTable + " " + 
							"FROM  " + p_strFVSOutputTable + " a, " + 
							p_strFVSOutputCycleYearTable + " b " + 
							"WHERE a.standid=b.standid AND " + 
							"a.year <> b.year1 AND a.year <> b.year2 AND a.year <> b.year3 AND a.year > b.year4 GROUP BY a.standid";




				}
				return "";
			}
            /// <summary>
			/// Returns the SQL used to insert the RX pretreatment year for each of the 4 cycles 
			/// </summary>
			/// <param name="p_strInsertTable"></param>
			/// <param name="p_strFVSOutputCycleYearTable"></param>
			/// <returns></returns>
			static public string FVSOutputTable_AuditInsertRxPreCycleYearTableSQL(string p_strInsertTable, string p_strFVSOutputCycleYearTable)
			{
				return "INSERT INTO " + p_strInsertTable + " (standid,pre_year1,pre_year2,pre_year3,pre_year4) " + 
				          "SELECT a.standid," + 
				          "IIF(a.year1_exists_yn='Y',a.year1,-1) AS pre_year1," + 
				          "IIF(a.year2_exists_yn='Y',a.year2,-1) AS pre_year2," + 
				          "IIF(a.year3_exists_yn='Y',a.year3,-1) AS pre_year3," + 
				          "IIF(a.year4_exists_yn='Y',a.year4,-1) AS pre_year4 " + 
				          "FROM " + p_strFVSOutputCycleYearTable + " a";
			}

			static public string FVSOutputTable_AuditUpdateRxPostCyleYearSQL(string p_strSourceTable, 
				string p_strJoinTable,
				string p_strCycle)
			{

				switch (p_strCycle)
				{
					case "1":
						return "UPDATE " + p_strSourceTable + " a " +
							"INNER JOIN " + p_strJoinTable + " b " + 
							"ON a.standid=b.standid SET a.post_year1=b.post_year1";
					case "2":
						return  "UPDATE " + p_strSourceTable + " a " +
							"INNER JOIN " + p_strJoinTable + " b " + 
							"ON a.standid=b.standid SET a.post_year2=IIF(a.pre_year2=-1,-1,b.post_year2)";
					case "3":
						return   "UPDATE " + p_strSourceTable + " a " +
							"INNER JOIN " + p_strJoinTable + " b " + 
							"ON a.standid=b.standid SET a.post_year3=IIF(a.pre_year3=-1,-1,b.post_year3)";
					case "4":
						return  "UPDATE " + p_strSourceTable + " a " +
							"INNER JOIN " + p_strJoinTable + " b " + 
							"ON a.standid=b.standid SET a.post_year4=IIF(a.pre_year4=-1,-1,b.post_year4)";




				}
				return "";
			}
            /// <summary>
            /// Audit to find missing stand,year combinations.
            /// Define the SQL Queries that will identify standid,year rows in the tree list tables that
            /// are not found in the FVS_SUMMARY table.
            /// </summary>
            /// <param name="p_strIntoTempTreeListTableName"></param>
            /// <param name="p_strIntoTempSummaryTableName"></param>
            /// <param name="p_strIntoTempMissingRowsTableName"></param>
            /// <param name="p_strTreeListTableName"></param>
            /// <param name="p_strSummaryTableName"></param>
            /// <returns></returns>
			static public string[] FVSOutputTable_AuditSelectTreeListCyleYearExistInFVSSummaryTableSQL(
                string p_strIntoTempTreeListTableName,
                string p_strIntoTempSummaryTableName,
                string p_strIntoTempMissingRowsTableName,
                string p_strTreeListTableName,
                string p_strSummaryTableName)
			{
                string[] strSQL = new string[6];
                strSQL[0] = "SELECT COUNT(*) AS TREECOUNT, STANDID,YEAR " +
                            "INTO " + p_strIntoTempTreeListTableName + " " +
                            "FROM " + p_strTreeListTableName + " " +
                            "WHERE standid IS NOT NULL " +
                            "GROUP BY CASEID,STANDID,YEAR";

                strSQL[1] = "SELECT DISTINCT STANDID,YEAR " +
                            "INTO " + p_strIntoTempSummaryTableName + " " +
                            "FROM " + p_strSummaryTableName + " " +
                            "WHERE standid IS NOT NULL";

                strSQL[2] = "SELECT a.STANDID,a.YEAR,a.TREECOUNT, " + 
                              "SUM(IIF(a.standid=b.standid AND a.year=b.year,1,0)) AS ROWCOUNT " + 
                            "INTO " + p_strIntoTempMissingRowsTableName + " " + 
                            "FROM " + p_strIntoTempTreeListTableName + " a," + 
                             "(SELECT DISTINCT STANDID,YEAR " + 
                              "FROM " + p_strIntoTempSummaryTableName + " " + 
                              "WHERE STANDID IS NOT NULL) b " + 
                            "WHERE a.standid=b.standid " + 
                            "GROUP BY a.standid,a.year,a.treecount";

                strSQL[3] = "DELETE FROM " + p_strIntoTempMissingRowsTableName + "  WHERE ROWCOUNT > 0";

                strSQL[4] = "DROP TABLE " + p_strIntoTempSummaryTableName;

                strSQL[5] = "DROP TABLE " + p_strIntoTempTreeListTableName;


                return strSQL;

               
			}
            /// <summary>
            /// View the PRE-POST records from the FVS Output table (FVS_STRCLASS) that will be retrieved 
            /// based on sequential number assignment filters and removal code filters
            /// </summary>
            /// <param name="p_strFVSOutTable">FVS_STRCLASS table name</param>
            /// <param name="p_strSeqNumMatrixTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <param name="p_oItem"></param>
            /// <returns></returns>
            static public string FVSOutputTable_StrClassPrePostSeqNumByCycle(string p_strFVSOutTable,string p_strSeqNumMatrixTable,FVSPrePostSeqNumItem p_oItem)
            {

                string strSQL="";
                string strPreRemovalCode = "";
                string strPostRemovalCode = "";
                for (int x=1;x<=4;x++)
                {
                    if (x == 1)
                    {
                        strPreRemovalCode = p_oItem.RxCycle1PreStrClassBeforeTreeRemovalYN == "Y" ? "0" : "1";
                        strPostRemovalCode = p_oItem.RxCycle1PostStrClassBeforeTreeRemovalYN == "Y" ? "0" : "1";
                    }
                    else if (x == 2)
                    {
                        strPreRemovalCode = p_oItem.RxCycle2PreStrClassBeforeTreeRemovalYN == "Y" ? "0" : "1";
                        strPostRemovalCode = p_oItem.RxCycle2PostStrClassBeforeTreeRemovalYN == "Y" ? "0" : "1";
                    }
                    else if (x == 3)
                    {
                        strPreRemovalCode = p_oItem.RxCycle3PreStrClassBeforeTreeRemovalYN == "Y" ? "0" : "1";
                        strPostRemovalCode = p_oItem.RxCycle3PostStrClassBeforeTreeRemovalYN == "Y" ? "0" : "1";
                    }
                    else if (x == 4)
                    {
                        strPreRemovalCode = p_oItem.RxCycle4PreStrClassBeforeTreeRemovalYN == "Y" ? "0" : "1";
                        strPostRemovalCode = p_oItem.RxCycle4PostStrClassBeforeTreeRemovalYN == "Y" ? "0" : "1";
                    }
                    strSQL = strSQL + "SELECT b.seqnum, b.cycle,b.type, a.* " +
                                      "FROM " + p_strFVSOutTable.Trim() + " AS a," + 
                                         "(SELECT seqnum,standid,year,'" + x.ToString().Trim() + "' AS cycle,'PRE' AS [type] " + 
                                          "FROM " + p_strSeqNumMatrixTable + " WHERE CYCLE" + x.ToString().Trim() + "_PRE_YN='Y')  AS b " + 
                                     "WHERE a.standid=b.standid AND a.year=b.year AND a.Removal_Code=" + strPreRemovalCode + " " + 
                                     "UNION " + 
                                     "SELECT b.seqnum, b.cycle,b.type, a.* " +
                                      "FROM " + p_strFVSOutTable + " AS a," +
                                         "(SELECT seqnum,standid,year,'" + x.ToString().Trim() + "' AS cycle,'POST' AS [type] " + 
                                          "FROM " + p_strSeqNumMatrixTable + " WHERE CYCLE" + x.ToString().Trim() + "_POST_YN='Y')  AS b " +
                                     "WHERE a.standid=b.standid AND a.year=b.year AND a.Removal_Code=" + strPostRemovalCode + " " + 
                                     "UNION ";

                }
                strSQL = strSQL.Substring(0, strSQL.Length - 6);
                return strSQL;
               


            }
            /// <summary>
            /// View the PRE-POST records from the FVS Output table that will be retrieved 
            /// based on sequential number assignment filters
            /// </summary>
            /// <param name="p_strFVSOutTable">FVS output table</param>
            /// <param name="p_strSeqNumMatrixTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <returns></returns>
            static public string FVSOutputTable_GenericPrePostSeqNumByCycle(string p_strFVSOutTable, string p_strSeqNumMatrixTable)
            {

                string strSQL = "";

                for (int x = 1; x <= 4; x++)
                {
                    strSQL = strSQL + "SELECT b.seqnum, b.cycle,b.type, a.* " +
                                      "FROM " + p_strFVSOutTable.Trim() + " AS a," +
                                         "(SELECT seqnum,standid,year,'" + x.ToString().Trim() + "' AS cycle,'PRE' AS [type] " +
                                          "FROM " + p_strSeqNumMatrixTable + " WHERE CYCLE" + x.ToString().Trim() + "_PRE_YN='Y')  AS b " +
                                     "WHERE a.standid=b.standid AND a.year=b.year " +
                                     "UNION " +
                                     "SELECT b.seqnum, b.cycle,b.type, a.* " +
                                      "FROM " + p_strFVSOutTable + " AS a," +
                                         "(SELECT seqnum,standid,year,'" + x.ToString().Trim() + "' AS cycle,'POST' AS [type] " +
                                          "FROM " + p_strSeqNumMatrixTable + " WHERE CYCLE" + x.ToString().Trim() + "_POST_YN='Y')  AS b " +
                                     "WHERE a.standid=b.standid AND a.year=b.year " +
                                     "UNION ";

                }
                strSQL = strSQL.Substring(0, strSQL.Length - 6);
                return strSQL;
            }
            
			/// <summary>
			/// make sure the FVSOut tables pre_year1,pre_year2,pre_year3,and pre_year4 cycle values are the same in the
			/// FVS Summary table
			/// </summary>
			/// <param name="p_strFVSOutCycleYearTable"></param>
			/// <param name="p_strSummaryCycleYearTable"></param>
			/// <returns></returns>
			static public string FVSOutputTable_AuditSelectFVSOutCycleYearInFVSSummaryCycleYearTableSQL(string p_strFVSOutPrePostCycleYearTable, string p_strSummaryPrePostCycleYearTable)
			{
				return "SELECT a.standid,a.pre_year1,a.pre_year2,a.pre_year3,a.pre_year4," + 
										"b.pre_year1 AS summary_pre_year1,b.pre_year2 AS summary_pre_year2," + 
										"b.pre_year3 AS summary_pre_year3,b.pre_year4 AS summary_pre_year4 " + 
					  "FROM " + p_strSummaryPrePostCycleYearTable + " a, " + p_strFVSOutPrePostCycleYearTable + " b " +  
					  "WHERE a.standid=b.standid AND (a.pre_year1 <> b.pre_year1 OR " + 
					        "a.pre_year2 <> b.pre_year2 OR " + 
					        "a.pre_year3 <> b.pre_year3 OR " +  
					        "a.pre_year4 <> b.pre_year4)";
			}

			/// <summary>
			/// Update the FVSOut pre/post year table from the values in the FVSOUT summary pre/post table.
			/// If we find the FVSOUT summary pre/post table year in the FVSOut table then set the value
			/// to the summary year, otherwise, set the value to -1
			/// </summary>
			/// <param name="p_strFVSOutPrePostTable"></param>
			/// <param name="p_strFVSOutTable"></param>
			/// <param name="p_strFVSOutPrePostSummaryTable"></param>
			/// <returns></returns>
			static public string FVSOutputTable_AuditSelectIntoPreYears(string p_strIntoTable,string p_strFVSOutTable, string p_strFVSOutPrePostSummaryTable,string p_strCycle)
			{
				if (p_strFVSOutTable.Trim().ToUpper().IndexOf("STRCLASS",0) < 0)
				{
					return "SELECT b.standid, a.pre_year" + p_strCycle.Trim() + " " +  
						"INTO " + p_strIntoTable + " " + 
						"FROM " + p_strFVSOutPrePostSummaryTable + " a," + 
						p_strFVSOutTable + " b " + 
						"WHERE b.standid=a.standid AND b.year=a.pre_year" + p_strCycle.Trim();
				}
				else
				{
					return "SELECT b.standid, a.pre_year" + p_strCycle.Trim() + " " +  
						"INTO " + p_strIntoTable + " " + 
						"FROM " + p_strFVSOutPrePostSummaryTable + " a," + 
						p_strFVSOutTable + " b " + 
						"WHERE b.standid=a.standid AND b.year=a.pre_year" + p_strCycle.Trim() + " AND " + 
						      "b.Removal_Code IS NOT NULL AND b.Removal_Code=0";
				}

			}
			static public string FVSOutputTable_AuditSelectIntoPostYears(string p_strIntoTable,string p_strFVSOutPrePostTable,string p_strFVSOutTable)
			{
				return "SELECT a.standid, IIF(a.pre_year1=-1,-1,y1.min_post_year1) AS post_year1," + 
										 "IIF(a.pre_year2=-1,-1,y2.min_post_year2) AS post_year2," + 
										 "IIF(a.pre_year3=-1,-1,y3.min_post_year3) AS post_year3," + 
										 "IIF(a.pre_year4=-1,-1,y4.min_post_year4) AS post_year4 " + 
					   "INTO " + p_strIntoTable + " " + 
					   "FROM " + p_strFVSOutPrePostTable + " a," + 
						"(SELECT MIN(a.year) AS min_post_year4,a.standid " + 
						"FROM " + p_strFVSOutTable + " a," +  p_strFVSOutPrePostTable + " b " +  
						"WHERE a.year > b.pre_year4  AND a.standid=b.standid " + 
						"GROUP BY a.standid) y4," +
						"(SELECT MIN(a.year) AS min_post_year3,a.standid " + 
						"FROM " + p_strFVSOutTable + " a," + p_strFVSOutPrePostTable + " b " + 
						"WHERE a.year > b.pre_year3  AND a.standid=b.standid " + 
						"GROUP BY a.standid) y3," +
						"(SELECT MIN(a.year) AS min_post_year2,a.standid " + 
						"FROM " + p_strFVSOutTable + " a," + p_strFVSOutPrePostTable + " b " +  
						"WHERE a.year > b.pre_year2  AND a.standid=b.standid " + 
						"GROUP BY a.standid) y2," + 
						"(SELECT MIN(a.year) AS min_post_year1,a.standid " + 
						"FROM " + p_strFVSOutTable + " a," + p_strFVSOutPrePostTable + " b " +
						"WHERE a.year > b.pre_year1 AND a.standid=b.standid " + 
						"GROUP BY a.standid) y1 " +
					"WHERE a.standid=y4.standid AND " + 
						  "a.standid=y3.standid AND " + 
						  "a.standid=y2.standid AND " + 
						  "a.standid=y1.standid";
			}
			static public string FVSOutputTable_AuditUpdatePostYears(string p_strFVSOutPrePostTable,string p_strFVSOutPostWorkTable)
			{
				return "UPDATE " + p_strFVSOutPrePostTable.Trim() +  " a " +
					   "INNER JOIN " + p_strFVSOutPostWorkTable.Trim() + " b " + 
					   "ON a.standid=b.standid " + 
					   "SET a.post_year1=b.post_year1,a.post_year2=b.post_year2, " + 
					       "a.post_year3=b.post_year3,a.post_year4=b.post_year4";
			}
            /// <summary>
            /// Every FIA tree in the FVS_CUTLIST table should be found in the tree table. This query
            /// formats the variant and treeid column into the FVS_TREE_ID. FVS created trees are 
            /// not included (Seedlings and Compacted).
            /// </summary>
            /// <param name="p_strFvsTreeIdAuditTable"></param>
            /// <param name="p_strCasesTable"></param>
            /// <param name="p_strCutListTable"></param>
            /// <param name="p_strFVSCutListPrePostSeqNumTable"></param>
            /// <param name="p_strRxPackage"></param>
            /// <param name="p_strRx"></param>
            /// <param name="p_strRxCycle"></param>
            /// <param name="p_strRxYear"></param>
            /// <returns></returns>
            static public string FVSOutputTable_AuditFVSTreeId(string p_strFvsTreeIdAuditTable,
                                                               string p_strCasesTable,
                                                               string p_strCutListTable,
                                                               string p_strFVSCutListPrePostSeqNumTable,
                                                               string p_strRxPackage,
                                                               string p_strRx,
                                                               string p_strRxCycle,
                                                               string p_strRxYear)
            {
                return "INSERT INTO " + p_strFvsTreeIdAuditTable + " " +
                                "(biosum_cond_id, rxpackage,rx,rxcycle,rxyear,fvs_variant, fvs_tree_id) " +
                                 "SELECT DISTINCT c.StandID AS biosum_cond_id,'" + p_strRxPackage.Trim() + "' AS rxpackage," +
                                "'" + p_strRx.Trim() + "' AS rx,'" + p_strRxCycle.Trim() + "' AS rxcycle," +
                                "CSTR(t.year) AS rxyear," +
                                "c.Variant AS fvs_variant, IIf(Len(Trim(t.treeid))=4," +
                                "c.variant+'000'+Trim(t.treeid),IIf(Len(Trim(t.treeid))=5," +
                                "c.variant+'00'+Trim(t.treeid),IIf(Len(Trim(t.treeid))=6," +
                                "c.variant+'0'+Trim(t.treeid),c.variant+Trim(t.treeid)))) AS fvs_tree_id " +
                                "FROM " + p_strCasesTable + " c," + p_strCutListTable + " t," + p_strFVSCutListPrePostSeqNumTable + " p " +
                                "WHERE c.CaseID = t.CaseID AND t.standid=p.standid AND t.year=p.year AND  " + 
                                      "p.cycle" + p_strRxCycle.Trim() + "_PRE_YN='Y' AND " + 
                                      "MID(t.treeid, 1, 2) <> 'ES'  AND MID(t.treeid, 1, 2)<> 'CM'";


  
            }
            public class VolumesAndBiomass
            {
                public VolumesAndBiomass()
                {
                }
                public static string FVSOut_BuiltInputTableForVolumeCalculation_Step1(string p_strInputVolumesTable,string p_strFvsTreeTable,string p_strRxPackage)
                {
                    string strColumns = "id,biosum_cond_id,invyr,fvs_variant,spcd,dbh,ht,actualht,cr,fvs_tree_id";
                    return "INSERT INTO " + p_strInputVolumesTable + " " +
                            "(" + strColumns + ") " + 
                            "SELECT id,biosum_cond_id,CINT(rxyear) AS invyr," + 
                                    "fvs_variant," + 
                                    "IIF(FvsCreatedTree_YN='Y',CINT(fvs_species),-1) AS spcd," +
                                    "dbh,estht,ht,pctcr,fvs_tree_id " +
                            "FROM " + p_strFvsTreeTable + " " +
                            "WHERE rxpackage='" + p_strRxPackage.Trim() + "' AND rxcycle<>'1'"; 
                }
                public static string FVSOut_BuiltInputTableForVolumeCalculation_Step1(string p_strInputVolumesTable, string p_strFvsTreeTable)
                {
                    string strColumns = "id,biosum_cond_id,invyr,fvs_variant,spcd,dbh,ht,actualht,cr,fvs_tree_id";
                    return "INSERT INTO " + p_strInputVolumesTable + " " +
                            "(" + strColumns + ") " +
                            "SELECT id,biosum_cond_id,CINT(rxyear) AS invyr," +
                                    "fvs_variant," +
                                    "IIF(FvsCreatedTree_YN='Y',CINT(fvs_species),-1) AS spcd," +
                                    "dbh,estht,ht,pctcr,fvs_tree_id " +
                            "FROM " + p_strFvsTreeTable + " " +
                            "WHERE rxcycle<>'1'";
                }
                public static string FVSOut_BuildInputTableForVolumeCalculation_Step2(
                                        string p_strInputVolumesTable,
                                        string p_strFIATreeTable,
                                        string p_strFIAPlotTable,
                                        string p_strFIACondTable)
                {
                    return "UPDATE " + p_strInputVolumesTable + " i " +
                           "INNER JOIN ((" + p_strFIATreeTable + " t " +
                                        "INNER JOIN " + p_strFIACondTable + " c " +
                                        "ON t.biosum_cond_id=c.biosum_cond_id) " +
                                        "INNER JOIN " + p_strFIAPlotTable + " p " +
                                        "ON p.biosum_plot_id=c.biosum_plot_id) " +
                           "ON i.fvs_tree_id=t.fvs_tree_id " +
                           "SET i.spcd=t.spcd," +
                               "i.statuscd=IIF(t.statuscd IS NULL,1,t.statuscd)," +
                               "i.treeclcd=t.treeclcd," +
                               "i.cull=IIF(t.cull IS NULL,0,t.cull)," +
                               "i.roughcull=IIF(t.roughcull IS NULL,0,t.roughcull)," +
                               "i.decaycd=IIF(t.decaycd IS NULL,0,t.decaycd)";
                }
                public static string FVSOut_BuildInputTableForVolumeCalculation_Step3(
                                        string p_strInputVolumesTable,
                                        string p_strFIACondTable)
                {
                    return "UPDATE " + p_strInputVolumesTable + " i " +
                                      "INNER JOIN " + p_strFIACondTable + " c " +
                                      "ON i.biosum_cond_id=c.biosum_cond_id " +
                                      "SET i.vol_loc_grp=IIF(INSTR(1,c.vol_loc_grp,'22') > 0,'S26LEOR',c.vol_loc_grp)," + 
                                          "i.statuscd=IIF(i.statuscd IS NULL,1,i.statuscd)," +
                                          "i.cull=IIF(i.cull IS NULL,0,i.cull)," + 
                                          "i.roughcull=IIF(i.roughcull IS NULL,0,i.roughcull)," +
                                          "i.decaycd=IIF(i.decaycd IS NULL,0,i.decaycd)";
                }
                /// <summary>
                /// Create and Populate a CULL work table with TOTALCULL (cull + roughcull)
                /// </summary>
                /// <returns></returns>
                public static string FVSOut_BuildInputTableForVolumeCalculation_Step4(
                                     string p_strCullTable, 
                                     string p_strInputVolumesTable)
                {
                    return "SELECT id, IIF(cull IS NOT NULL AND roughcull IS NOT NULL, cull + roughcull," +
                                      "IIF(cull IS NOT NULL,cull," +
                                      "IIF(roughcull IS NOT NULL, roughcull,0))) AS totalcull " +
                           "INTO " + p_strCullTable + " " + 
                           "FROM " + p_strInputVolumesTable;

                }
                public static string FIAPlotInput_BuildInputTableForVolumeCalculation_Step1(string p_strInputVolumesTable,string p_strFIATreeTable,string p_strColumns,string p_strValues)
                {
                    return  "INSERT INTO " + p_strInputVolumesTable + " " +
                            "(" + p_strColumns + ") SELECT " + p_strValues + " FROM " + p_strFIATreeTable;
                }
                public static string FIAPlotInput_BuildInputTableForVolumeCalculation_Step2(
                                        string p_strInputVolumesTable,
                                        string p_strFIATreeTable,
                                        string p_strFIAPlotTable,
                                        string p_strFIACondTable)
                {
                    return "UPDATE " + p_strInputVolumesTable + " i " +
                           "INNER JOIN ((" + p_strFIATreeTable + " t " +
                                        "INNER JOIN " + p_strFIACondTable + " c " +
                                        "ON t.biosum_cond_id=c.biosum_cond_id) " +
                                        "INNER JOIN " + p_strFIAPlotTable + " p " +
                                        "ON p.biosum_plot_id=c.biosum_plot_id) " +
                           "ON i.tre_cn=t.cn " +
                           "SET i.spcd=t.spcd," +
                               "i.statuscd=IIF(t.statuscd IS NULL,1,t.statuscd)," +
                               "i.treeclcd=t.treeclcd," +
                               "i.cull=IIF(t.cull IS NULL,0,t.cull)," +
                               "i.roughcull=IIF(t.roughcull IS NULL,0,t.roughcull)," +
                               "i.decaycd=IIF(t.decaycd IS NULL,0,t.decaycd)";
                }
                public static string FIAPlotInput_BuildInputTableForVolumeCalculation_Step3(
                                        string p_strInputVolumesTable,
                                        string p_strFIACondTable)
                {
                    return "UPDATE " + p_strInputVolumesTable + " i " +
                                      "INNER JOIN " + p_strFIACondTable + " c " +
                                      "ON i.cnd_cn=c.biosum_cond_id " +
                                      "SET i.vol_loc_grp=IIF(INSTR(1,c.vol_loc_grp,'22') > 0,'S26LEOR',c.vol_loc_grp)," +
                                          "i.statuscd=IIF(i.statuscd IS NULL,1,i.statuscd)," +
                                          "i.cull=IIF(i.cull IS NULL,0,i.cull)," +
                                          "i.roughcull=IIF(i.roughcull IS NULL,0,i.roughcull)," +
                                          "i.decaycd=IIF(i.decaycd IS NULL,0,i.decaycd)";
                }
                public static string FIAPlotInput_BuildInputTableForVolumeCalculation_Step4(
                                    string p_strCullTable,
                                    string p_strInputVolumesTable)
                {
                    return "SELECT tre_cn, IIF(cull IS NOT NULL AND roughcull IS NOT NULL, cull + roughcull," +
                                      "IIF(cull IS NOT NULL,cull," +
                                      "IIF(roughcull IS NOT NULL, roughcull,0))) AS totalcull " +
                           "INTO " + p_strCullTable + " " +
                           "FROM " + p_strInputVolumesTable;

                }
               //public static string FIAPlotInput_BuildInputTableForVolumeCalculation_Step2(string p_strInputVolumesTable, string p_strFIACondTable)
               // {
               //     return "UPDATE " + p_strInputVolumesTable + " f INNER JOIN " + p_strFIACondTable + " c ON f.CND_CN = c.BIOSUM_COND_ID SET f.vol_loc_grp=IIF(INSTR(1,c.vol_loc_grp,'22') > 0,'S26LEOR',c.vol_loc_grp)";
               // }

                public class PNWRS
                {
                    public PNWRS()
                    {
                    }
                    /// <summary>
                    /// Update the TREECLCD column for all trees 
                    /// by comparing SPCD,ROUGHCULL,STATUSCD and TOTALCULL columns
                    /// </summary>
                    /// <param name="p_strCullTable"></param>
                    /// <param name="p_strInputVolumesTable"></param>
                    /// <returns></returns>
                    public static string FVSOut_BuildInputTableForVolumeCalculation_Step5(
                                         string p_strCullTable,
                                         string p_strInputVolumesTable)
                    {
                        return "UPDATE " + p_strInputVolumesTable + " a " +
                                          "INNER JOIN " + p_strCullTable + " b " +
                                          "ON a.id=b.id " +
                                          "SET a.treeclcd=" +
                                          "IIF(a.SpCd IN (62,65,66,106,133,138,304,321,322,475,756,758,990),3," +
                                          "IIF(a.StatusCd=2,3," +
                                          "IIF(b.totalcull < 75,2," +
                                          "IIF(a.roughcull > 37.5,3,4))))";
                    }
                    /// <summary>
                    /// Update the TREECLCD column for TREECLCD=3 and dead trees 
                    /// by comparing SPCD,DBH,STATUSCD and DECAYCD columns
                    /// </summary>
                    /// <param name="p_strCullTable"></param>
                    /// <param name="p_strInputVolumesTable"></param>
                    /// <returns></returns>
                    public static string FVSOut_BuildInputTableForVolumeCalculation_Step6(
                                         string p_strCullTable,
                                         string p_strInputVolumesTable)
                    {
                        return "UPDATE " + p_strInputVolumesTable + " a " +
                                          "INNER JOIN " + p_strCullTable + " b " +
                                          "ON a.id=b.id " +
                                          "SET a.treeclcd=" +
                                          "IIF(a.DecayCd > 1,4,IIF(a.dbh < 9 AND a.SpCd < 300,4,a.treeclcd)) " +
                                          "WHERE a.treeclcd=3 AND a.statuscd=2 AND a.SpCd NOT IN (62,65,66,106,133,138,304,321,322,475,756,758,990)";
                    }
                    /// <summary>
                    /// Update the TREECLCD column for all trees 
                    /// by comparing SPCD,ROUGHCULL,STATUSCD and TOTALCULL columns
                    /// </summary>
                    /// <param name="p_strCullTable"></param>
                    /// <param name="p_strInputVolumesTable"></param>
                    /// <returns></returns>
                    public static string FIAPlotInput_BuildInputTableForVolumeCalculation_Step5(
                                         string p_strCullTable,
                                         string p_strInputVolumesTable)
                    {
                        return "UPDATE " + p_strInputVolumesTable + " a " +
                                          "INNER JOIN " + p_strCullTable + " b " +
                                          "ON a.tre_cn=b.tre_cn " +
                                          "SET a.treeclcd=" +
                                          "IIF(a.SpCd IN (62,65,66,106,133,138,304,321,322,475,756,758,990),3," +
                                          "IIF(a.StatusCd=2,3," +
                                          "IIF(b.totalcull < 75,2," +
                                          "IIF(a.roughcull > 37.5,3,4))))";
                    }
                    /// <summary>
                    /// Update the TREECLCD column for TREECLCD=3 and dead trees 
                    /// by comparing SPCD,DBH,STATUSCD and DECAYCD columns
                    /// </summary>
                    /// <param name="p_strCullTable"></param>
                    /// <param name="p_strInputVolumesTable"></param>
                    /// <returns></returns>
                    public static string FIAPlotInput_BuildInputTableForVolumeCalculation_Step6(
                                         string p_strCullTable,
                                         string p_strInputVolumesTable)
                    {
                        return "UPDATE " + p_strInputVolumesTable + " a " +
                                          "INNER JOIN " + p_strCullTable + " b " +
                                          "ON a.tre_cn=b.tre_cn " +
                                          "SET a.treeclcd=" +
                                          "IIF(a.DecayCd > 1,4,IIF(a.DIA < 9 AND a.SpCd < 300,4,a.treeclcd)) " +
                                          "WHERE a.treeclcd=3 AND a.statuscd=2 AND a.SpCd NOT IN (62,65,66,106,133,138,304,321,322,475,756,758,990)";
                    }
                }
                /// <summary>
                /// Insert into the MS Access Biosum Volume table
                /// the formatted data in the input volumes table.
                /// This extra step is needed before importing to 
                /// Oracle because the performance of formatting of data from Access to 
                /// the Oracle Linked table is slow.
                /// </summary>
                /// <param name="p_strInputVolumesTable"></param>
                /// <param name="p_strOracleBiosumVolumesTable"></param>
                /// <returns></returns>
                public static string FVSOut_BuildInputTableForVolumeCalculation_Step7(
                            string p_strInputVolumesTable,
                            string p_strBiosumVolumesTable)
                {
                    string strColumns = "STATECD,COUNTYCD,PLOT,INVYR,VOL_LOC_GRP,TREE,SPCD,DIA,HT," +
                                        "ACTUALHT,CR,STATUSCD,TREECLCD,ROUGHCULL,CULL,TRE_CN,CND_CN,PLT_CN";
                      
                    string strValues = "CINT(MID(BIOSUM_COND_ID,6,2)) AS STATECD,CINT(MID(BIOSUM_COND_ID,12,3)) AS COUNTYCD,CINT(MID(BIOSUM_COND_ID,16,5)) AS PLOT," +
                                       "INVYR,VOL_LOC_GRP,ID AS TREE,SPCD,DBH AS DIA,HT,ACTUALHT,CR,STATUSCD,TREECLCD,ROUGHCULL,CULL," +
                                       "CSTR(ID) AS TRE_CN,BIOSUM_COND_ID AS CND_CN,MID(BIOSUM_COND_ID,1,LEN(BIOSUM_COND_ID)-1) AS PLT_CN";

                    return "INSERT INTO " + p_strBiosumVolumesTable + " " +
                           "(" + strColumns + ") " + 
                           "SELECT " + strValues + " " + 
                           "FROM " + p_strInputVolumesTable;
                }
                /// <summary>
                /// Insert records into the fcs oracle linked table
                /// </summary>
                /// <param name="p_strBiosumVolumesTable"></param>
                /// <param name=""></param>
                /// <returns></returns>
                public static string FVSOut_BuildInputTableForVolumeCalculation_Step8(
                            string p_strBiosumVolumesTable,
                            string p_strOracleFCSBiosumVolumesLinkedTable)
                {
                    string strColumns = "STATECD,COUNTYCD,PLOT,INVYR,VOL_LOC_GRP,TREE,SPCD,DIA,HT," +
                                        "ACTUALHT,CR,STATUSCD,TREECLCD,ROUGHCULL,CULL,TRE_CN,CND_CN,PLT_CN";

                    
                    return "INSERT INTO " + p_strOracleFCSBiosumVolumesLinkedTable + " " +
                           "(" + strColumns + ") " +
                           "SELECT " + strColumns + " " +
                           "FROM " + p_strBiosumVolumesTable;
                }
                /// <summary>
                /// Update the FVS_TREE table with the volumes and biomass
                /// values that Oracle returned
                /// </summary>
                /// <param name="p_strFvsTreeTable"></param>
                /// <param name="p_strOracleBiosumVolumesTable"></param>
                /// <returns></returns>
                public static string FVSOut_BuildInputTableForVolumeCalculation_Step9(
                              string p_strFvsTreeTable,
                              string p_strOracleBiosumVolumesTable)
                {
                    return "UPDATE " + p_strFvsTreeTable + " f " +
                           "INNER JOIN " + p_strOracleBiosumVolumesTable + " o " +
                           "ON f.id = o.tree " +
                           "SET f.volcsgrs=o.VOLCSGRS_CALC," +
                               "f.volcfgrs=o.VOLCFGRS_CALC," +
                               "f.volcfnet=o.VOLCFNET_CALC," +
                               "f.drybiot=o.DRYBIOT_CALC," +
                               "f.drybiom=o.DRYBIOM_CALC," + 
                               "f.voltsgrs=o.VOLTSGRS_CALC" ;
                }

               
                


                
            }

			
			

		}
		public class TravelTime
		{
            public string m_strTravelTimeTable;
            private bool _bLoadDataSources = true;
            private Queries _oQueries=null;
			public TravelTime()
			{
			}
            public Queries ReferenceQueries
			{
				get {return _oQueries;}
				set {_oQueries=value;}
			}
            public bool LoadDatasource
            {
                get { return _bLoadDataSources; }
                set { _bLoadDataSources = value; }
            }
			

            public void LoadDatasources()
            {
                m_strTravelTimeTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("PLOT");


                if (this.m_strTravelTimeTable.Trim().Length == 0)
                {

                    MessageBox.Show("!!Could Not Locate Travel Times Table!!", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    ReferenceQueries.m_intError = -1;
                    return;
                }
            }
		}
		public class Audit
		{
			

			public Audit()
			{
			}

		}
		public class FIAPlot
		{
			public string m_strPlotTable;
			public string m_strTreeTable;
			public string m_strCondTable;
			public string m_strSiteIndexTable;
			public string m_strTreeRegionalBiomassTable;
			public string m_strPopEvalTable;
			public string m_strPopEstnUnitTable;
			public string m_strPopStratumTable;
			public string m_strPPSATable;
            public string m_strBiosumPopStratumAdjustmentFactorsTable;


			private Queries _oQueries=null;	
			private bool _bLoadDataSources=true;
			string m_strSql="";

			public FIAPlot()
			{
			}

			public Queries ReferenceQueries
			{
				get {return _oQueries;}
				set {_oQueries=value;}
			}
			public bool LoadDatasource
			{
				get {return _bLoadDataSources;}
				set {_bLoadDataSources=value;}
			}


			public void LoadDatasources()
			{
				m_strPlotTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("PLOT");
				m_strTreeTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREE");
				m_strCondTable= ReferenceQueries.m_oDataSource.getValidDataSourceTableName("CONDITION");
				
			
				if (this.m_strPlotTable.Trim().Length == 0) 
				{
					
					MessageBox.Show("!!Could Not Locate Plot Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
				if (this.m_strCondTable.Trim().Length == 0)
				{
					MessageBox.Show("!!Could Not Locate Condition Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
				if (this.m_strTreeTable.Trim().Length == 0 && this._oQueries._strScenarioType!="core")
				{
					MessageBox.Show("!!Could Not Locate Tree Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
				
			
			}
            
            static public string[] FIADBPlotInput_CalculateAdjustmentFactorsSQL(
                string p_strPpsaTable,
                string p_strPopEstUnitTable,
                string p_strPopStratumTable,
                string p_strPopEvalTable,
                string p_strFIADBPlotTable,
                string p_strFIADBCondTable,
                string p_strRsCd,
                string p_strEvalId,
                string p_strCondProp)
            {
                string[] strSQL = new string[21];

                //
                //CREATE THE BIOSUM PLOT TABLE WITH 
                //ONLY THE CURRENTLY SELECTED EVALUATION
                //
                strSQL[0] = "SELECT p.* " + 
                            "INTO BIOSUM_PLOT " + 
                            "FROM " + p_strFIADBPlotTable + " p, " + 
                                "(SELECT DISTINCT PLT_CN " + 
                                 "FROM " + p_strPpsaTable + " " + 
                                 "WHERE RSCD=" + p_strRsCd + " AND EVALID=" + p_strEvalId + ") ppsa " + 
                            "WHERE p.CN = ppsa.PLT_CN";
                //
                //CREATE BIOSUM_PPSA
                //
                strSQL[1] = "SELECT ppsa.*,p.CYCLE,p.SUBCYCLE " + 
                            "INTO BIOSUM_PPSA " + 
                            "FROM " + p_strPpsaTable + " ppsa " + 
                            "INNER JOIN BIOSUM_PLOT p " + 
                            "ON ppsa.PLT_CN=p.CN " + 
                            "WHERE RSCD=" + p_strRsCd + " AND EVALID=" + p_strEvalId;
                //
                //CREATE BIOSUM COND TABLE
                //
                strSQL[2] = "SELECT c.* INTO BIOSUM_COND FROM " + p_strFIADBCondTable + " c," + 
                            "(SELECT CN FROM BIOSUM_PLOT) p " + 
                            "WHERE c.PLT_CN = p.CN";

                //change hazardous condition to sampled
                strSQL[3] = "UPDATE BIOSUM_COND SET cond_status_cd = 1 " +
                            "WHERE COND_NONSAMPLE_REASN_CD = 5";

                //update condition satatus to NONSAMPLED if the condition proportion is less than .25
                strSQL[4] = "UPDATE BIOSUM_COND SET cond_status_cd = 5 " +
                            "WHERE cond_status_cd = 1 AND condprop_unadj < ." + p_strCondProp;

                //join pop_estn_unit,pop_stratum,pop_eval tables into biosum_eus_temp
                strSQL[5] = "SELECT pe.rscd, pe.evalid,ps.estn_unit,ps.stratumcd," +
                                   "pe.eval_descr,peu.estn_unit_descr,peu.arealand_eu," +
                                   "peu.areatot_eu , ps.p1pointcnt, ps.p2pointcnt," +
                                   "peu.p1pntcnt_eu as p1pointcnt_eu,peu.area_used," +
                                   "ps.adj_factor_macr,ps.adj_factor_subp," +
                                   "ps.adj_factor_micr,ps.expns,pe.LAND_ONLY " +
                            "INTO BIOSUM_EUS_TEMP " +
                            "FROM " + p_strPopEstUnitTable + " peu," +
                                      p_strPopEvalTable + " pe," + p_strPopStratumTable + " ps " +
                                      "WHERE  ((pe.rscd=" + p_strRsCd + " AND pe.EVALID=" + p_strEvalId + ")) AND " +
                                              "(pe.rscd = ps.rscd AND pe.evalid = ps.evalid) AND " +
                                              "(ps.rscd = peu.rscd AND ps.evalid = peu.evalid AND " +
                                               "ps.estn_unit = peu.estn_unit)";

                //
                //SUM UP UNADJUSTED FACTORS FOR DENIED ACCESS
                //
                strSQL[6] = "SELECT DISTINCT ppsa.evalid, ppsa.estn_unit, ppsa.statecd, ppsa.stratumcd, ppsa.plot," +
                                            "ppsa.countycd, ppsa.subcycle, ppsa.cycle, ppsa.unitcd," +
                                            "SUM(IIF(eus.LAND_ONLY='N'," +
                                                    "IIF(c.COND_STATUS_CD IN (1,2,3,4),0,c.MACRPROP_UNADJ)," +
                                                    "IIF(c.COND_STATUS_CD IN (1,2,3),0,c.MACRPROP_UNADJ))) as denied_macr," +
                                            "SUM(IIF(eus.LAND_ONLY='N'," +
                                                    "IIF(c.COND_STATUS_CD IN (1,2,3,4),0,c.MICRPROP_UNADJ)," +
                                                    "IIF(c.COND_STATUS_CD IN (1,2,3),0,c.MICRPROP_UNADJ))) as denied_micr," +
                                            "SUM(IIF(eus.LAND_ONLY='N'," +
                                                    "IIF(c.COND_STATUS_CD IN (1,2,3,4),0,c.SUBPPROP_UNADJ)," +
                                                    "IIF(c.COND_STATUS_CD IN (1,2,3),0,c.SUBPPROP_UNADJ))) as denied_subp," +
                                            "SUM(IIF(eus.LAND_ONLY='N',IIF(c.COND_STATUS_CD IN (1,2,3,4),0,c.CONDPROP_UNADJ),IIF(c.COND_STATUS_CD IN (1,2,3),0,c.CONDPROP_UNADJ))) as denied_cond " +
                            "INTO BIOSUM_PPSA_DENIED_ACCESS " +
                            "FROM BIOSUM_PPSA ppsa," +
                                 "BIOSUM_COND c," +
                                 "BIOSUM_EUS_TEMP eus " +
                            "WHERE (ppsa.plt_cn = c.plt_cn) AND " +
                                  "(ppsa.rscd = eus.rscd AND " +
                                   "ppsa.evalid = eus.evalid AND " +
                                   "ppsa.stratumcd = eus.stratumcd AND " +
                                   "ppsa.estn_unit = eus.estn_unit) " +
                            "GROUP BY ppsa.evalid, ppsa.estn_unit, ppsa.statecd, ppsa.stratumcd," +
                                     "ppsa.plot, ppsa.countycd, ppsa.subcycle, ppsa.cycle, ppsa.unitcd";

                //DELETE DENIED_COND=1
                strSQL[7] = "DELETE FROM BIOSUM_PPSA_DENIED_ACCESS WHERE DENIED_COND =  1";
                //JOIN THE 2 TABLES
                strSQL[8] = "SELECT ppsa.*," +
                                   "denied.denied_macr," +
                                   "denied.denied_micr," +
                                   "denied.denied_subp," +
                                   "denied.denied_cond " +
                            "INTO BIOSUM_PPSA_TEMP " +
                            "FROM BIOSUM_PPSA_DENIED_ACCESS denied," +
                                 "BIOSUM_PPSA ppsa " +
                           "WHERE ppsa.evalid = denied.evalid AND " +
                                 "ppsa.estn_unit = denied.estn_unit AND " +
                                 "ppsa.stratumcd = denied.stratumcd AND " +
                                 "ppsa.plot = denied.plot AND " +
                                 "ppsa.statecd = denied.statecd AND " +
                                 "ppsa.countycd  = denied.countycd AND " +
                                 "ppsa.subcycle  = denied.subcycle AND " +
                                 "ppsa.cycle     = denied.cycle AND " +
                                 "ppsa.unitcd    = denied.unitcd";
                //
                //CALCULATE ADJUSTMENTS
                //
                strSQL[9] = "SELECT DISTINCT " +
                                   "eus.rscd, eus.evalid, eus.estn_unit, eus.stratumcd," +
                                   "eus.arealand_eu, eus.areatot_eu," +
                                   "eus.area_used, rowcount.p2pointcnt_man as p2pointcnt_man," +
                                   "eus.p1pointcnt, eus.p1pointcnt_eu, eus.p2pointcnt," +
                                   "SUM(c.MACRPROP_UNADJ * " +
                                        "IIF(eus.LAND_ONLY='N'," +
                                        "IIF(c.COND_STATUS_CD IN (1,2,3,4),1,0)," +
                                        "IIF(c.COND_STATUS_CD IN (1,2,3),1,0))) / " +
                                            "SUM(c.macrprop_unadj) as pmh_macr," +
                                   "SUM(c.MICRPROP_UNADJ * " +
                                       "IIF(eus.LAND_ONLY='N'," +
                                       "IIF(c.COND_STATUS_CD IN (1,2,3,4),1,0)," +
                                       "IIF(c.COND_STATUS_CD IN (1,2,3),1,0))) / " +
                                           "SUM(c.micrprop_unadj) as pmh_micr," +
                                   "SUM(c.SUBPPROP_UNADJ * " +
                                       "IIF(eus.LAND_ONLY='N'," +
                                       "IIF(c.COND_STATUS_CD IN (1,2,3,4),1,0)," +
                                       "IIF(c.COND_STATUS_CD IN (1,2,3),1,0))) / " +
                                           "SUM(c.subpprop_unadj) as pmh_sub," +
                                   "SUM(c.CONDPROP_UNADJ * " +
                                       "IIF(eus.LAND_ONLY='N'," +
                                       "IIF(c.COND_STATUS_CD IN (1,2,3,4),1,0)," +
                                       "IIF(c.COND_STATUS_CD IN (1,2,3),1,0))) / " +
                                           "SUM(c.condprop_unadj) as pmh_cond " +
                            "INTO BIOSUM_EUS_ACCESS " +
                            "FROM BIOSUM_COND c," +
                                 "BIOSUM_PPSA_TEMP ppsa," +
                                 "BIOSUM_EUS_TEMP  eus," +
                                "(SELECT eus_count.rscd," +
                                        "eus_count.evalid," +
                                        "eus_count.estn_unit," +
                                        "eus_count.stratumcd," +
                                        "COUNT(ppsa_count.PLT_CN) as p2pointcnt_man " +
                                 "FROM BIOSUM_PPSA_TEMP ppsa_count," +
                                      "BIOSUM_EUS_TEMP eus_count " +
                                 "WHERE (ppsa_count.rscd = eus_count.rscd AND " +
                                        "ppsa_count.evalid = eus_count.evalid AND " +
                                        "ppsa_count.stratumcd = eus_count.stratumcd AND " +
                                        "ppsa_count.estn_unit = eus_count.estn_unit) " +
                                 "GROUP BY eus_count.rscd," +
                                          "eus_count.evalid," +
                                          "eus_count.estn_unit," +
                                          "eus_count.stratumcd " +
                                ") AS ROWCOUNT " +
                            "WHERE (ppsa.plt_cn = c.plt_cn) AND " +
                                  "(ppsa.rscd = eus.rscd AND " +
                                   "ppsa.evalid = eus.evalid AND " +
                                   "ppsa.stratumcd = eus.stratumcd AND " +
                                   "ppsa.estn_unit = eus.estn_unit) AND " +
                                  "(eus.rscd = rowcount.rscd AND " +
                                   "eus.evalid = rowcount.evalid AND " +
                                   "eus.estn_unit = rowcount.estn_unit AND " +
                                   "eus.stratumcd = rowcount.stratumcd) " +
                           "GROUP BY eus.rscd," +
                                    "eus.evalid," +
                                    "eus.estn_unit," +
                                    "eus.stratumcd," +
                                    "eus.arealand_eu," +
                                    "eus.areatot_eu," +
                                    "eus.area_used," +
                                    "eus.p1pointcnt," +
                                    "eus.p1pointcnt_eu," +
                                    "eus.p2pointcnt," +
                                    "p2pointcnt_man";

                strSQL[10] = "ALTER TABLE BIOSUM_EUS_ACCESS ADD COLUMN STRATUM_AREA DOUBLE";
                strSQL[11] = "ALTER TABLE BIOSUM_EUS_ACCESS ADD COLUMN DOUBLE_SAMPLING INTEGER";
                //
                //CALCULATE STRATUM AREA
                //
                strSQL[12] = "UPDATE BIOSUM_EUS_ACCESS " +
                             "SET double_sampling = " +
                                 "IIF(p1pointcnt_eu is null OR " +
                                     "p1pointcnt is null OR " +
                                     "p1pointcnt=p1pointcnt_eu," +
                                     "0,1)," +
                            "stratum_area = " +
                                "IIF(p1pointcnt_eu is NOT null AND " +
                                    "p1pointcnt_eu > 0," +
                                    "area_used*p1pointcnt/p1pointcnt_eu,0)";
                //
                //MERGE BIOSUM_EUS_ACCESS WITH BIOSUM_EUS_TEMP INTO  biosum_pop_stratum_adjustment_factors
                //
                strSQL[13] = "SELECT a.rscd,a.evalid," +
                                  "a.estn_unit,a.stratumcd," +
                                  "a.p2pointcnt_man,a.stratum_area," +
                                  "a.double_sampling," +
                                  "a.pmh_macr,a.pmh_micr," +
                                  "a.pmh_sub,a.pmh_cond," +
                                  "b.eval_descr,b.estn_unit_descr," +
                                  "b.adj_factor_macr, b.adj_factor_subp," +
                                  "b.adj_factor_micr, b.expns " +
                          "INTO biosum_pop_stratum_adjustment_factors " +
                          "FROM BIOSUM_EUS_ACCESS a " +
                          "INNER JOIN BIOSUM_EUS_TEMP b " +
                          "ON a.rscd = b.rscd AND " +
                             "a.evalid = b.evalid AND " +
                             "a.estn_unit = b.estn_unit AND " +
                             "a.stratumcd = b.stratumcd";

                strSQL[14] = "ALTER TABLE biosum_pop_stratum_adjustment_factors ADD COLUMN stratum_cn CHAR(34)";
                //
                //UPDATE THE  biosum_pop_stratum_adjustment_factors TABLE 
                //WITH THE KEY COLUMN FROM THE POP_STRATUM TABLE
                //
                strSQL[15] = "UPDATE biosum_pop_stratum_adjustment_factors b  " + 
                             "INNER JOIN " + p_strPopStratumTable + " ps " + 
                             "ON ps.RSCD = b.RSCD AND ps.EVALID=b.EVALID AND " + 
                                "ps.ESTN_UNIT=b.ESTN_UNIT AND ps.STRATUMCD = b.STRATUMCD " + 
                             "SET b.STRATUM_CN=ps.CN " + 
                             "WHERE b.RSCD=" + p_strRsCd + " AND b.EVALID=" + p_strEvalId;
                //
                //CLEAN UP
                //
                strSQL[16] = "DROP TABLE BIOSUM_PPSA";
                strSQL[17] = "DROP TABLE BIOSUM_EUS_TEMP";
                strSQL[18] = "DROP TABLE BIOSUM_PPSA_DENIED_ACCESS";
                strSQL[19] = "DROP TABLE BIOSUM_PPSA_TEMP";
                strSQL[20] = "DROP TABLE BIOSUM_EUS_ACCESS";
                return strSQL;
            }
             
			
		}
		public class Processor
		{
            private Queries _oQueries = null;
            private bool _bLoadDataSources = true;
            public string m_strAdditionalHarvestCostsTable;
			public Processor()
			{
			}
            public Queries ReferenceQueries
            {
                get { return _oQueries; }
                set { _oQueries = value; }
            }
            public bool LoadDatasource
            {
                get { return _bLoadDataSources; }
                set { _bLoadDataSources = value; }
            }
            public void LoadDatasources()
            {
              
            }
           

			public static string AuditFvsOut_SelectIntoUnionOfFVSTreeTables(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strIntoTable,RxPackageItem_Collection p_oRxPackageItem_Collection, string[] p_strFVSVariantsArray,string p_strColumnList)
			{
               
				string strSql="SELECT DISTINCT " + p_strColumnList +  " " + 
					          "INTO " + p_strIntoTable + " " + 
					          "FROM (";
				int x,y;
				for (x=0;x<=p_strFVSVariantsArray.Length-1;x++)
				{
                    for (y = 0; y <= p_oRxPackageItem_Collection.Count - 1; y++)
                    {
                        if (p_oAdo.TableExist(p_oConn, "fvs_tree_IN_" + p_strFVSVariantsArray[x].Trim() + "_P" + p_oRxPackageItem_Collection.Item(y).RxPackageId + "_TREE_CUTLIST"))
                        {
                            strSql = strSql + "SELECT " + p_strColumnList + " " +
                                              "FROM fvs_tree_IN_" + p_strFVSVariantsArray[x].Trim() + "_P" + p_oRxPackageItem_Collection.Item(y).RxPackageId + "_TREE_CUTLIST UNION ";
                        }
                    }
				}
				if (strSql.IndexOf("UNION",0) > 0) strSql = strSql.Substring(0,strSql.Length - 6) + ")";
               
				return strSql;
				
			}
            public static List<string> AuditFvsOut_SelectIntoUnionOfFVSTreeTablesUsingListArray(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strIntoTable, RxPackageItem_Collection p_oRxPackageItem_Collection, string[] p_strFVSVariantsArray, string p_strColumnList)
            {
                List<string> strList = new List<string>();

                string strTable = "";

                string strSql = "";
                int x, y;
                for (x = 0; x <= p_strFVSVariantsArray.Length - 1; x++)
                {
                    for (y = 0; y <= p_oRxPackageItem_Collection.Count - 1; y++)
                    {
                        
                        if (p_oAdo.TableExist(p_oConn, "fvs_tree_IN_" + p_strFVSVariantsArray[x].Trim() + "_P" + p_oRxPackageItem_Collection.Item(y).RxPackageId + "_TREE_CUTLIST"))
                        {
                            if (strTable.Trim().Length > 0)
                            {
                                strTable = "fvs_tree_IN_" + p_strFVSVariantsArray[x].Trim() + "_P" + p_oRxPackageItem_Collection.Item(y).RxPackageId + "_TREE_CUTLIST";
                                strSql = "INSERT INTO " + p_strIntoTable + " (" + p_strColumnList + ") " +
                                         "SELECT DISTINCT " + p_strColumnList + " " +
                                         "FROM " + strTable;

                               
                            }
                            else
                            {

                                strTable = "fvs_tree_IN_" + p_strFVSVariantsArray[x].Trim() + "_P" + p_oRxPackageItem_Collection.Item(y).RxPackageId + "_TREE_CUTLIST";

                                strSql = "SELECT DISTINCT " + p_strColumnList + " " +
                                         "INTO " + p_strIntoTable + " " +
                                         "FROM " + strTable;
                               
                            }
                            strList.Add(strSql);
                            
                        }
                    }
                }
                return strList;

            }
			
		}
        public class ProcessorScenarioRun
        {
            private string strSQL = "";
            private string _strScenarioId = "";
            public string ScenarioId
            {
                get { return _strScenarioId; }
                set { _strScenarioId = value; }
            }
            public ProcessorScenarioRun()
            {
            }
            public static string InitializeBinsTable(string p_strTreeBinsTableName,
                                                     string p_strClass)
            {
                return "UPDATE " + p_strTreeBinsTableName + " b " +
                        "SET b." + p_strClass + "_Util_Logs_count = 0," +
                        "b." + p_strClass + "_Util_Logs_TPA = 0," +
                        "b." + p_strClass + "_Util_Chips_count = 0," +
                        "b." + p_strClass + "_Util_Chips_TPA = 0," +
                        "b." + p_strClass + "_Util_Logs_merch_wt = 0," +
                        "b." + p_strClass + "_Util_Logs_merch_vol = 0," +
                        "b." + p_strClass + "_Util_Logs_biomass_wt = 0," +
                        "b." + p_strClass + "_Util_Logs_biomass_vol = 0," +
                        "b." + p_strClass + "_Util_Chips_merch_wt = 0," +
                        "b." + p_strClass + "_Util_Chips_merch_vol = 0," +
                        "b." + p_strClass + "_Util_Chips_biomass_wt = 0," +
                        "b." + p_strClass + "_Util_Chips_biomass_vol = 0," +
                        "b." + p_strClass + "_NonUtil_Logs_count = 0," +
                        "b." + p_strClass + "_NonUtil_Logs_TPA = 0," +
                        "b." + p_strClass + "_NonUtil_Chips_count = 0," +
                        "b." + p_strClass + "_NonUtil_Chips_TPA = 0," +
                        "b." + p_strClass + "_NonUtil_Logs_merch_wt = 0," +
                        "b." + p_strClass + "_NonUtil_Logs_merch_vol = 0," +
                        "b." + p_strClass + "_NonUtil_Logs_biomass_wt = 0," +
                        "b." + p_strClass + "_NonUtil_Logs_biomass_vol = 0," +
                        "b." + p_strClass + "_NonUtil_Chips_merch_wt = 0," +
                        "b." + p_strClass + "_NonUtil_Chips_merch_vol = 0," +
                        "b." + p_strClass + "_NonUtil_Chips_biomass_wt = 0," +
                        "b." + p_strClass + "_NonUtil_Chips_biomass_vol = 0";



            }
            public static string UpdateBinsTable(string p_strDiametersTableName,
                                          string p_strTreeBinsTableName,
                                          string p_strTreeDataInTableName,
                                          string p_strClass)
            {
                return "UPDATE " + p_strDiametersTableName + " d " +
                        "INNER JOIN (" + p_strTreeDataInTableName + " t " +
                        "INNER JOIN " + p_strTreeBinsTableName + " b " +
                        "ON  (t.rxpackage = b.rxpackage AND " +
                             "t.rx = b.rx AND " +
                             "t.rxcycle = b.rxcycle AND " +
                             "t.fvs_tree_id = b.fvs_tree_id)) " +
                        "ON d.DBH = b.DBH AND d.species_group=b.species_group " +
                        "SET b." + p_strClass + "_Util_Logs_count = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Logs_Bl='Y',1,0)," +
                        "b." + p_strClass + "_Util_Logs_TPA = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Logs_Bl='Y',t.TPA,0)," +
                        "b." + p_strClass + "_Util_Chips_count = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Chips_Bl='Y',1,0)," +
                        "b." + p_strClass + "_Util_Chips_TPA = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Chips_Bl='Y',t.TPA,0)," +
                        "b." + p_strClass + "_Util_Logs_merch_wt = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Logs_Bl='Y',t.[MERCH_WT_GT],0)," +
                        "b." + p_strClass + "_Util_Logs_merch_vol = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Logs_Bl='Y',t.[MERCH_VOL_CF],0)," +
                        "b." + p_strClass + "_Util_Logs_biomass_wt = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Logs_Bl='Y',t.[CHIP_WT_GT],0)," +
                        "b." + p_strClass + "_Util_Logs_biomass_vol = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Logs_Bl='Y',t.[CHIP_VOL_CF],0)," +
                        "b." + p_strClass + "_Util_Chips_merch_wt = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Chips_Bl='Y',t.[MERCH_WT_GT],0)," +
                        "b." + p_strClass + "_Util_Chips_merch_vol = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Chips_Bl='Y',t.[MERCH_VOL_CF],0)," +
                        "b." + p_strClass + "_Util_Chips_biomass_wt = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Chips_Bl='Y',t.[CHIP_WT_GT],0)," +
                        "b." + p_strClass + "_Util_Chips_biomass_vol = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Chips_Bl='Y',t.[CHIP_VOL_CF],0)," + 
                        "b." + p_strClass + "_NonUtil_Logs_count = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Logs_Bl='Y',1,0)," +
                        "b." + p_strClass + "_NonUtil_Logs_TPA = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Logs_Bl='Y',t.TPA,0)," +
                        "b." + p_strClass + "_NonUtil_Chips_count = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Chips_Bl='Y',1,0)," +
                        "b." + p_strClass + "_NonUtil_Chips_TPA = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Chips_Bl='Y',t.TPA,0)," +
                        "b." + p_strClass + "_NonUtil_Logs_merch_wt = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Logs_Bl='Y',t.[MERCH_WT_GT],0)," +
                        "b." + p_strClass + "_NonUtil_Logs_merch_vol = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Logs_Bl='Y',t.[MERCH_VOL_CF],0)," +
                        "b." + p_strClass + "_NonUtil_Logs_biomass_wt = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Logs_Bl='Y',t.[CHIP_WT_GT],0)," +
                        "b." + p_strClass + "_NonUtil_Logs_biomass_vol = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Logs_Bl='Y',t.[CHIP_VOL_CF],0)," +
                        "b." + p_strClass + "_NonUtil_Chips_merch_wt = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Chips_Bl='Y',t.[MERCH_WT_GT],0)," +
                        "b." + p_strClass + "_NonUtil_Chips_merch_vol = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Chips_Bl='Y',t.[MERCH_VOL_CF],0)," +
                        "b." + p_strClass + "_NonUtil_Chips_biomass_wt = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Chips_Bl='Y',t.[CHIP_WT_GT],0)," +
                        "b." + p_strClass + "_NonUtil_Chips_biomass_vol = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Chips_Bl='Y',t.[CHIP_VOL_CF],0) ";
                


            }
            public static string InitializeHwdBinsTable(string p_strTreeBinsTableName,
                                                     string p_strClass)
            {
                return "UPDATE " + p_strTreeBinsTableName + " b " +
                        "SET b.HWD_" + p_strClass + "_Util_Logs_count = 0," +
                        "b.HWD_" + p_strClass + "_Util_Logs_TPA = 0," +
                        "b.HWD_" + p_strClass + "_Util_Chips_count = 0," +
                        "b.HWD_" + p_strClass + "_Util_Chips_TPA = 0," +
                        "b.HWD_" + p_strClass + "_Util_Logs_merch_wt = 0," +
                        "b.HWD_" + p_strClass + "_Util_Logs_merch_vol = 0," +
                        "b.HWD_" + p_strClass + "_Util_Logs_biomass_wt = 0," +
                        "b.HWD_" + p_strClass + "_Util_Logs_biomass_vol = 0," +
                        "b.HWD_" + p_strClass + "_Util_Chips_merch_wt = 0," +
                        "b.HWD_" + p_strClass + "_Util_Chips_merch_vol = 0," +
                        "b.HWD_" + p_strClass + "_Util_Chips_biomass_wt = 0," +
                        "b.HWD_" + p_strClass + "_Util_Chips_biomass_vol = 0," + 
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_count = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_TPA = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_count = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_TPA = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_merch_wt = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_merch_vol = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_biomass_wt = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_biomass_vol = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_merch_wt = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_merch_vol = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_biomass_wt = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_biomass_vol = 0";



            }
            public static string UpdateHwdBinsTable(string p_strDiametersTableName,
                                          string p_strTreeHwdBinsTableName,
                                          string p_strTreeDataInTableName,
                                          string p_strClass)
            {
                return "UPDATE " + p_strDiametersTableName + " d " +
                        "INNER JOIN (" + p_strTreeDataInTableName + " t " +
                        "INNER JOIN " + p_strTreeHwdBinsTableName + " b " +
                        "ON  (t.rxpackage = b.rxpackage AND " +
                             "t.rx = b.rx AND " +
                             "t.rxcycle = b.rxcycle AND " +
                             "t.fvs_tree_id = b.fvs_tree_id)) " +
                        "ON d.DBH = b.DBH AND d.species_group=b.species_group " +
                        "SET b.HWD_" + p_strClass + "_Util_Logs_count = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Logs_Bl='Y',1,0)," +
                        "b.HWD_" + p_strClass + "_Util_Logs_TPA = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Logs_Bl='Y',t.TPA,0)," +
                        "b.HWD_" + p_strClass + "_Util_Chips_count = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Chips_Bl='Y',1,0)," +
                        "b.HWD_" + p_strClass + "_Util_Chips_TPA = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Chips_Bl='Y',t.TPA,0)," +
                        "b.HWD_" + p_strClass + "_Util_Logs_merch_wt = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Logs_Bl='Y',t.[MERCH_WT_GT],0)," +
                        "b.HWD_" + p_strClass + "_Util_Logs_merch_vol = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Logs_Bl='Y',t.[MERCH_VOL_CF],0)," +
                        "b.HWD_" + p_strClass + "_Util_Logs_biomass_wt = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Logs_Bl='Y',t.[CHIP_WT_GT],0)," +
                        "b.HWD_" + p_strClass + "_Util_Logs_biomass_vol = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Logs_Bl='Y',t.[CHIP_VOL_CF],0)," +
                        "b.HWD_" + p_strClass + "_Util_Chips_merch_wt = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Chips_Bl='Y',t.[MERCH_WT_GT],0)," +
                        "b.HWD_" + p_strClass + "_Util_Chips_merch_vol = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Chips_Bl='Y',t.[MERCH_VOL_CF],0)," +
                        "b.HWD_" + p_strClass + "_Util_Chips_biomass_wt = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Chips_Bl='Y',t.[CHIP_WT_GT],0)," +
                        "b.HWD_" + p_strClass + "_Util_Chips_biomass_vol = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.Util_Chips_Bl='Y',t.[CHIP_VOL_CF],0), " +
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_count = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Logs_Bl='Y',1,0)," +
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_TPA = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Logs_Bl='Y',t.TPA,0)," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_count = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Chips_Bl='Y',1,0)," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_TPA = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Chips_Bl='Y',t.TPA,0)," +
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_merch_wt = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Logs_Bl='Y',t.[MERCH_WT_GT],0)," +
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_merch_vol = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Logs_Bl='Y',t.[MERCH_VOL_CF],0)," +
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_biomass_wt = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Logs_Bl='Y',t.[CHIP_WT_GT],0)," +
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_biomass_vol = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Logs_Bl='Y',t.[CHIP_VOL_CF],0)," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_merch_wt = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Chips_Bl='Y',t.[MERCH_WT_GT],0)," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_merch_vol = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Chips_Bl='Y',t.[MERCH_VOL_CF],0)," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_biomass_wt = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Chips_Bl='Y',t.[CHIP_WT_GT],0)," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_biomass_vol = " +
                              "IIF(TRIM(d.class)='" + p_strClass + "' And d.NonUtil_Chips_Bl='Y',t.[CHIP_VOL_CF],0) " + 
                    "WHERE t.spcd > 299";
                       


            }
            public static string SumBinByPlotRxSpcGrpDbhGrp(string p_strFromTableName, string p_strIntoTableName)
            {
                return "SELECT BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE, species_group, diam_group," +
                        " Sum(BC_Util_Logs_count) AS BC_Util_Logs_count," +
                        " Sum(BC_Util_Logs_TPA) AS BC_Util_Logs_TPA," +
                        " Sum(BC_Util_Logs_merch_wt) AS BC_Util_Logs_merch_wt," +
                        " Sum(BC_Util_Logs_merch_vol) AS BC_Util_Logs_merch_vol," +
                        " Sum(BC_Util_Logs_biomass_vol) AS BC_Util_Logs_biomass_vol," +
                        " Sum(BC_Util_Logs_biomass_wt) AS BC_Util_Logs_biomass_wt," +
                        " Sum(BC_Util_Chips_count) AS BC_Util_Chips_count," +
                        " Sum(BC_Util_Chips_TPA) AS BC_Util_Chips_TPA," +
                        " Sum(BC_Util_Chips_merch_wt) AS BC_Util_Chips_merch_wt," +
                        " Sum(BC_Util_Chips_merch_vol) AS BC_Util_Chips_merch_vol," +
                        " Sum(BC_Util_Chips_biomass_wt) AS BC_Util_Chips_biomass_wt," +
                        " Sum(BC_Util_Chips_biomass_vol) AS BC_Util_Chips_biomass_vol," +
                        " Sum(BC_NonUtil_Logs_count) AS BC_NonUtil_Logs_count," +
                        " Sum(BC_NonUtil_Logs_TPA) AS BC_NonUtil_Logs_TPA," +
                        " Sum(BC_NonUtil_Logs_merch_wt) AS BC_NonUtil_Logs_merch_wt," +
                        " Sum(BC_NonUtil_Logs_merch_vol) AS BC_NonUtil_Logs_merch_vol," +
                        " Sum(BC_NonUtil_Logs_biomass_vol) AS BC_NonUtil_Logs_biomass_vol," +
                        " Sum(BC_NonUtil_Logs_biomass_wt) AS BC_NonUtil_Logs_biomass_wt," +
                        " Sum(BC_NonUtil_Chips_count) AS BC_NonUtil_Chips_count," +
                        " Sum(BC_NonUtil_Chips_TPA) AS BC_NonUtil_Chips_TPA," +
                        " Sum(BC_NonUtil_Chips_merch_wt) AS BC_NonUtil_Chips_merch_wt," +
                        " Sum(BC_NonUtil_Chips_merch_vol) AS BC_NonUtil_Chips_merch_vol," +
                        " Sum(BC_NonUtil_Chips_biomass_wt) AS BC_NonUtil_Chips_biomass_wt," +
                        " Sum(BC_NonUtil_Chips_biomass_vol) AS BC_NonUtil_Chips_biomass_vol," +
                        " Sum(CHIPS_Util_Logs_count) AS CHIPS_Util_Logs_count," +
                        " Sum(CHIPS_Util_Logs_TPA) AS CHIPS_Util_Logs_TPA," +
                        " Sum(CHIPS_Util_Logs_merch_wt) AS CHIPS_Util_Logs_merch_wt," +
                        " Sum(CHIPS_Util_Logs_merch_vol) AS CHIPS_Util_Logs_merch_vol," +
                        " Sum(CHIPS_Util_Chips_count) AS CHIPS_Util_Chips_count," +
                        " Sum(CHIPS_Util_Chips_TPA) AS CHIPS_Util_Chips_TPA," +
                        " Sum(CHIPS_Util_Chips_merch_wt) AS CHIPS_Util_Chips_merch_wt," +
                        " Sum(CHIPS_Util_Chips_merch_vol) AS CHIPS_Util_Chips_merch_vol," +
                        " Sum(CHIPS_Util_Logs_biomass_wt) AS CHIPS_Util_Logs_biomass_wt," +
                        " Sum(CHIPS_Util_Logs_biomass_vol) AS CHIPS_Util_Logs_biomass_vol," +
                        " Sum(CHIPS_Util_Chips_biomass_wt) AS CHIPS_Util_Chips_biomass_wt," +
                        " Sum(CHIPS_Util_Chips_biomass_vol) AS CHIPS_Util_Chips_biomass_vol," +
                		" Sum(CHIPS_NonUtil_Logs_count) AS CHIPS_NonUtil_Logs_count," +
                        " Sum(CHIPS_NonUtil_Logs_TPA) AS CHIPS_NonUtil_Logs_TPA," +
                        " Sum(CHIPS_NonUtil_Logs_merch_wt) AS CHIPS_NonUtil_Logs_merch_wt," +
                        " Sum(CHIPS_NonUtil_Logs_merch_vol) AS CHIPS_NonUtil_Logs_merch_vol," +
                        " Sum(CHIPS_NonUtil_Chips_count) AS CHIPS_NonUtil_Chips_count," +
                        " Sum(CHIPS_NonUtil_Chips_TPA) AS CHIPS_NonUtil_Chips_TPA," +
                        " Sum(CHIPS_NonUtil_Chips_merch_wt) AS CHIPS_NonUtil_Chips_merch_wt," +
                        " Sum(CHIPS_NonUtil_Chips_merch_vol) AS CHIPS_NonUtil_Chips_merch_vol," +
                        " Sum(CHIPS_NonUtil_Logs_biomass_wt) AS CHIPS_NonUtil_Logs_biomass_wt," +
                        " Sum(CHIPS_NonUtil_Logs_biomass_vol) AS CHIPS_NonUtil_Logs_biomass_vol," +
                        " Sum(CHIPS_NonUtil_Chips_biomass_wt) AS CHIPS_NonUtil_Chips_biomass_wt," +
                        " Sum(CHIPS_NonUtil_Chips_biomass_vol) AS CHIPS_NonUtil_Chips_biomass_vol," +
                        " Sum(SMLOGS_Util_Logs_count) AS SMLOGS_Util_Logs_count," +
                        " Sum(SMLOGS_Util_Logs_TPA) AS SMLOGS_Util_Logs_TPA," +
                        " Sum(SMLOGS_Util_Logs_merch_wt) AS SMLOGS_Util_Logs_merch_wt," +
                        " Sum(SMLOGS_Util_Logs_merch_vol) AS SMLOGS_Util_Logs_merch_vol," +
                        " Sum(SMLOGS_Util_Logs_biomass_wt) AS SMLOGS_Util_Logs_biomass_wt," +
                        " Sum(SMLOGS_Util_Logs_biomass_vol) AS SMLOGS_Util_Logs_biomass_vol," +
                        " Sum(SMLOGS_Util_Chips_count) AS SMLOGS_Util_Chips_count," +
                        " Sum(SMLOGS_Util_Chips_TPA) AS SMLOGS_Util_Chips_TPA," +
                        " Sum(SMLOGS_Util_Chips_merch_wt) AS SMLOGS_Util_Chips_merch_wt," +
                        " Sum(SMLOGS_Util_Chips_merch_vol) AS SMLOGS_Util_Chips_merch_vol," +
                        " Sum(SMLOGS_Util_Chips_biomass_wt) AS SMLOGS_Util_Chips_biomass_wt," +
                        " Sum(SMLOGS_Util_Chips_biomass_vol) AS SMLOGS_Util_Chips_biomass_vol," +
                        " Sum(SMLOGS_NonUtil_Logs_count) AS SMLOGS_NonUtil_Logs_count," +
                        " Sum(SMLOGS_NonUtil_Logs_TPA) AS SMLOGS_NonUtil_Logs_TPA," +
                        " Sum(SMLOGS_NonUtil_Logs_merch_wt) AS SMLOGS_NonUtil_Logs_merch_wt," +
                        " Sum(SMLOGS_NonUtil_Logs_merch_vol) AS SMLOGS_NonUtil_Logs_merch_vol," +
                        " Sum(SMLOGS_NonUtil_Logs_biomass_wt) AS SMLOGS_NonUtil_Logs_biomass_wt," +
                        " Sum(SMLOGS_NonUtil_Logs_biomass_vol) AS SMLOGS_NonUtil_Logs_biomass_vol," +
                        " Sum(SMLOGS_NonUtil_Chips_count) AS SMLOGS_NonUtil_Chips_count," +
                        " Sum(SMLOGS_NonUtil_Chips_TPA) AS SMLOGS_NonUtil_Chips_TPA," +
                        " Sum(SMLOGS_NonUtil_Chips_merch_wt) AS SMLOGS_NonUtil_Chips_merch_wt," +
                        " Sum(SMLOGS_NonUtil_Chips_merch_vol) AS SMLOGS_NonUtil_Chips_merch_vol," +
                        " Sum(SMLOGS_NonUtil_Chips_biomass_wt) AS SMLOGS_NonUtil_Chips_biomass_wt," +
                        " Sum(SMLOGS_NonUtil_Chips_biomass_vol) AS SMLOGS_NonUtil_Chips_biomass_vol," +
                        " Sum(LGLOGS_Util_Logs_count) AS LGLOGS_Util_Logs_count," +
                        " Sum(LGLOGS_Util_Logs_TPA) AS LGLOGS_Util_Logs_TPA," +
                        " Sum(LGLOGS_Util_Logs_merch_wt) AS LGLOGS_Util_Logs_merch_wt," +
                        " Sum(LGLOGS_Util_Logs_merch_vol) AS LGLOGS_Util_Logs_merch_vol," +
                        " Sum(LGLOGS_Util_Logs_biomass_wt) AS LGLOGS_Util_Logs_biomass_wt," +
                        " Sum(LGLOGS_Util_Logs_biomass_vol) AS LGLOGS_Util_Logs_biomass_vol," +
                        " Sum(LGLOGS_Util_Chips_count) AS LGLOGS_Util_Chips_count," +
                        " Sum(LGLOGS_Util_Chips_TPA) AS LGLOGS_Util_Chips_TPA," +
                        " Sum(LGLOGS_Util_Chips_merch_wt) AS LGLOGS_Util_Chips_merch_wt," +
                        " Sum(LGLOGS_Util_Chips_merch_vol) AS LGLOGS_Util_Chips_merch_vol, " +
                        " Sum(LGLOGS_Util_Chips_biomass_wt) AS LGLOGS_Util_Chips_biomass_wt," +
                        " Sum(LGLOGS_Util_Chips_biomass_vol) AS LGLOGS_Util_Chips_biomass_vol," +
                     	" Sum(LGLOGS_NonUtil_Logs_count) AS LGLOGS_NonUtil_Logs_count," +
                        " Sum(LGLOGS_NonUtil_Logs_TPA) AS LGLOGS_NonUtil_Logs_TPA," +
                        " Sum(LGLOGS_NonUtil_Logs_merch_wt) AS LGLOGS_NonUtil_Logs_merch_wt," +
                        " Sum(LGLOGS_NonUtil_Logs_merch_vol) AS LGLOGS_NonUtil_Logs_merch_vol," +
                        " Sum(LGLOGS_NonUtil_Logs_biomass_wt) AS LGLOGS_NonUtil_Logs_biomass_wt," +
                        " Sum(LGLOGS_NonUtil_Logs_biomass_vol) AS LGLOGS_NonUtil_Logs_biomass_vol," +
                        " Sum(LGLOGS_NonUtil_Chips_count) AS LGLOGS_NonUtil_Chips_count," +
                        " Sum(LGLOGS_NonUtil_Chips_TPA) AS LGLOGS_NonUtil_Chips_TPA," +
                        " Sum(LGLOGS_NonUtil_Chips_merch_wt) AS LGLOGS_NonUtil_Chips_merch_wt," +
                        " Sum(LGLOGS_NonUtil_Chips_merch_vol) AS LGLOGS_NonUtil_Chips_merch_vol, " +
                        " Sum(LGLOGS_NonUtil_Chips_biomass_wt) AS LGLOGS_NonUtil_Chips_biomass_wt," +
                        " Sum(LGLOGS_NonUtil_Chips_biomass_vol) AS LGLOGS_NonUtil_Chips_biomass_vol " +
                        " INTO " + p_strIntoTableName + " " +
                        " FROM " + p_strFromTableName + " " +
                        " GROUP BY BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE, species_group, diam_group" +
                        " HAVING (((species_group) Is Not Null))";
            }
            public static string SumHwdBinByPlotRxSpcGrpDbhGrp(string p_strFromTableName, string p_strIntoTableName)
            {
                return "SELECT BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE, species_group, diam_group," +
                        " Sum(HWD_BC_Util_Logs_count) AS HWD_BC_Util_Logs_count," +
                        " Sum(HWD_BC_Util_Logs_TPA) AS HWD_BC_Util_Logs_TPA," +
                        " Sum(HWD_BC_Util_Logs_merch_wt) AS HWD_BC_Util_Logs_merch_wt," +
                        " Sum(HWD_BC_Util_Logs_merch_vol) AS HWD_BC_Util_Logs_merch_vol," +
                        " Sum(HWD_BC_Util_Logs_biomass_vol) AS HWD_BC_Util_Logs_biomass_vol," +
                        " Sum(HWD_BC_Util_Logs_biomass_wt) AS HWD_BC_Util_Logs_biomass_wt," +
                        " Sum(HWD_BC_Util_Chips_count) AS HWD_BC_Util_Chips_count," +
                        " Sum(HWD_BC_Util_Chips_TPA) AS HWD_BC_Util_Chips_TPA," +
                        " Sum(HWD_BC_Util_Chips_merch_wt) AS HWD_BC_Util_Chips_merch_wt," +
                        " Sum(HWD_BC_Util_Chips_merch_vol) AS HWD_BC_Util_Chips_merch_vol," +
                        " Sum(HWD_BC_Util_Chips_biomass_wt) AS HWD_BC_Util_Chips_biomass_wt," +
                        " Sum(HWD_BC_Util_Chips_biomass_vol) AS HWD_BC_Util_Chips_biomass_vol," +
                        " Sum(HWD_BC_NonUtil_Logs_count) AS HWD_BC_NonUtil_Logs_count," +
                        " Sum(HWD_BC_NonUtil_Logs_TPA) AS HWD_BC_NonUtil_Logs_TPA," +
                        " Sum(HWD_BC_NonUtil_Logs_merch_wt) AS HWD_BC_NonUtil_Logs_merch_wt," +
                        " Sum(HWD_BC_NonUtil_Logs_merch_vol) AS HWD_BC_NonUtil_Logs_merch_vol," +
                        " Sum(HWD_BC_NonUtil_Logs_biomass_vol) AS HWD_BC_NonUtil_Logs_biomass_vol," +
                        " Sum(HWD_BC_NonUtil_Logs_biomass_wt) AS HWD_BC_NonUtil_Logs_biomass_wt," +
                        " Sum(HWD_BC_NonUtil_Chips_count) AS HWD_BC_NonUtil_Chips_count," +
                        " Sum(HWD_BC_NonUtil_Chips_TPA) AS HWD_BC_NonUtil_Chips_TPA," +
                        " Sum(HWD_BC_NonUtil_Chips_merch_wt) AS HWD_BC_NonUtil_Chips_merch_wt," +
                        " Sum(HWD_BC_NonUtil_Chips_merch_vol) AS HWD_BC_NonUtil_Chips_merch_vol," +
                        " Sum(HWD_BC_NonUtil_Chips_biomass_wt) AS HWD_BC_NonUtil_Chips_biomass_wt," +
                        " Sum(HWD_BC_NonUtil_Chips_biomass_vol) AS HWD_BC_NonUtil_Chips_biomass_vol," +
                        " Sum(HWD_CHIPS_Util_Logs_count) AS HWD_CHIPS_Util_Logs_count," +
                        " Sum(HWD_CHIPS_Util_Logs_TPA) AS HWD_CHIPS_Util_Logs_TPA," +
                        " Sum(HWD_CHIPS_Util_Logs_merch_wt) AS HWD_CHIPS_Util_Logs_merch_wt," +
                        " Sum(HWD_CHIPS_Util_Logs_merch_vol) AS HWD_CHIPS_Util_Logs_merch_vol," +
                        " Sum(HWD_CHIPS_Util_Chips_count) AS HWD_CHIPS_Util_Chips_count," +
                        " Sum(HWD_CHIPS_Util_Chips_TPA) AS HWD_CHIPS_Util_Chips_TPA," +
                        " Sum(HWD_CHIPS_Util_Chips_merch_wt) AS HWD_CHIPS_Util_Chips_merch_wt," +
                        " Sum(HWD_CHIPS_Util_Chips_merch_vol) AS HWD_CHIPS_Util_Chips_merch_vol," +
                        " Sum(HWD_CHIPS_Util_Logs_biomass_wt) AS HWD_CHIPS_Util_Logs_biomass_wt," +
                        " Sum(HWD_CHIPS_Util_Logs_biomass_vol) AS HWD_CHIPS_Util_Logs_biomass_vol," +
                        " Sum(HWD_CHIPS_Util_Chips_biomass_wt) AS HWD_CHIPS_Util_Chips_biomass_wt," +
                        " Sum(HWD_CHIPS_Util_Chips_biomass_vol) AS HWD_CHIPS_Util_Chips_biomass_vol," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_count) AS HWD_CHIPS_NonUtil_Logs_count," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_TPA) AS HWD_CHIPS_NonUtil_Logs_TPA," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_merch_wt) AS HWD_CHIPS_NonUtil_Logs_merch_wt," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_merch_vol) AS HWD_CHIPS_NonUtil_Logs_merch_vol," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_count) AS HWD_CHIPS_NonUtil_Chips_count," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_TPA) AS HWD_CHIPS_NonUtil_Chips_TPA," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_merch_wt) AS HWD_CHIPS_NonUtil_Chips_merch_wt," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_merch_vol) AS HWD_CHIPS_NonUtil_Chips_merch_vol," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_biomass_wt) AS HWD_CHIPS_NonUtil_Logs_biomass_wt," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_biomass_vol) AS HWD_CHIPS_NonUtil_Logs_biomass_vol," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_biomass_wt) AS HWD_CHIPS_NonUtil_Chips_biomass_wt," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_biomass_vol) AS HWD_CHIPS_NonUtil_Chips_biomass_vol," +
                        " Sum(HWD_SMLOGS_Util_Logs_count) AS HWD_SMLOGS_Util_Logs_count," +
                        " Sum(HWD_SMLOGS_Util_Logs_TPA) AS HWD_SMLOGS_Util_Logs_TPA," +
                        " Sum(HWD_SMLOGS_Util_Logs_merch_wt) AS HWD_SMLOGS_Util_Logs_merch_wt," +
                        " Sum(HWD_SMLOGS_Util_Logs_merch_vol) AS HWD_SMLOGS_Util_Logs_merch_vol," +
                        " Sum(HWD_SMLOGS_Util_Logs_biomass_wt) AS HWD_SMLOGS_Util_Logs_biomass_wt," +
                        " Sum(HWD_SMLOGS_Util_Logs_biomass_vol) AS HWD_SMLOGS_Util_Logs_biomass_vol," +
                        " Sum(HWD_SMLOGS_Util_Chips_count) AS HWD_SMLOGS_Util_Chips_count," +
                        " Sum(HWD_SMLOGS_Util_Chips_TPA) AS HWD_SMLOGS_Util_Chips_TPA," +
                        " Sum(HWD_SMLOGS_Util_Chips_merch_wt) AS HWD_SMLOGS_Util_Chips_merch_wt," +
                        " Sum(HWD_SMLOGS_Util_Chips_merch_vol) AS HWD_SMLOGS_Util_Chips_merch_vol," +
                        " Sum(HWD_SMLOGS_Util_Chips_biomass_wt) AS HWD_SMLOGS_Util_Chips_biomass_wt," +
                        " Sum(HWD_SMLOGS_Util_Chips_biomass_vol) AS HWD_SMLOGS_Util_Chips_biomass_vol," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_count) AS HWD_SMLOGS_NonUtil_Logs_count," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_TPA) AS HWD_SMLOGS_NonUtil_Logs_TPA," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_merch_wt) AS HWD_SMLOGS_NonUtil_Logs_merch_wt," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_merch_vol) AS HWD_SMLOGS_NonUtil_Logs_merch_vol," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_biomass_wt) AS HWD_SMLOGS_NonUtil_Logs_biomass_wt," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_biomass_vol) AS HWD_SMLOGS_NonUtil_Logs_biomass_vol," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_count) AS HWD_SMLOGS_NonUtil_Chips_count," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_TPA) AS HWD_SMLOGS_NonUtil_Chips_TPA," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_merch_wt) AS HWD_SMLOGS_NonUtil_Chips_merch_wt," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_merch_vol) AS HWD_SMLOGS_NonUtil_Chips_merch_vol," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_biomass_wt) AS HWD_SMLOGS_NonUtil_Chips_biomass_wt," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_biomass_vol) AS HWD_SMLOGS_NonUtil_Chips_biomass_vol," +
                        " Sum(HWD_LGLOGS_Util_Logs_count) AS HWD_LGLOGS_Util_Logs_count," +
                        " Sum(HWD_LGLOGS_Util_Logs_TPA) AS HWD_LGLOGS_Util_Logs_TPA," +
                        " Sum(HWD_LGLOGS_Util_Logs_merch_wt) AS HWD_LGLOGS_Util_Logs_merch_wt," +
                        " Sum(HWD_LGLOGS_Util_Logs_merch_vol) AS HWD_LGLOGS_Util_Logs_merch_vol," +
                        " Sum(HWD_LGLOGS_Util_Logs_biomass_wt) AS HWD_LGLOGS_Util_Logs_biomass_wt," +
                        " Sum(HWD_LGLOGS_Util_Logs_biomass_vol) AS HWD_LGLOGS_Util_Logs_biomass_vol," +
                        " Sum(HWD_LGLOGS_Util_Chips_count) AS HWD_LGLOGS_Util_Chips_count," +
                        " Sum(HWD_LGLOGS_Util_Chips_TPA) AS HWD_LGLOGS_Util_Chips_TPA," +
                        " Sum(HWD_LGLOGS_Util_Chips_merch_wt) AS HWD_LGLOGS_Util_Chips_merch_wt," +
                        " Sum(HWD_LGLOGS_Util_Chips_merch_vol) AS HWD_LGLOGS_Util_Chips_merch_vol," +
                        " Sum(HWD_LGLOGS_Util_Chips_biomass_wt) AS HWD_LGLOGS_Util_Chips_biomass_wt," +
                        " Sum(HWD_LGLOGS_Util_Chips_biomass_vol) AS HWD_LGLOGS_Util_Chips_biomass_vol," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_count) AS HWD_LGLOGS_NonUtil_Logs_count," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_TPA) AS HWD_LGLOGS_NonUtil_Logs_TPA," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_merch_wt) AS HWD_LGLOGS_NonUtil_Logs_merch_wt," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_merch_vol) AS HWD_LGLOGS_NonUtil_Logs_merch_vol," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_biomass_wt) AS HWD_LGLOGS_NonUtil_Logs_biomass_wt," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_biomass_vol) AS HWD_LGLOGS_NonUtil_Logs_biomass_vol," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_count) AS HWD_LGLOGS_NonUtil_Chips_count," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_TPA) AS HWD_LGLOGS_NonUtil_Chips_TPA," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_merch_wt) AS HWD_LGLOGS_NonUtil_Chips_merch_wt," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_merch_vol) AS HWD_LGLOGS_NonUtil_Chips_merch_vol," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_biomass_wt) AS HWD_LGLOGS_NonUtil_Chips_biomass_wt," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_biomass_vol) AS HWD_LGLOGS_NonUtil_Chips_biomass_vol " +
                        " INTO " + p_strIntoTableName + " " +
                        " FROM " + p_strFromTableName + " " +
                        " GROUP BY BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE, species_group, diam_group" +
                        " HAVING (((species_group) Is Not Null))";
            }
            public static string SumBinByPlotRx(string p_strFromTableName, string p_strIntoTableName)
            {
                return "SELECT BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE," +
                        " Sum(BC_Util_Logs_count) AS BC_Util_Logs_count," +
                        " Sum(BC_Util_Logs_TPA) AS BC_Util_Logs_TPA," +
                        " Sum(BC_Util_Logs_merch_wt) AS BC_Util_Logs_merch_wt," +
                        " Sum(BC_Util_Logs_merch_vol) AS BC_Util_Logs_merch_vol," +
                        " Sum(BC_Util_Logs_biomass_vol) AS BC_Util_Logs_biomass_vol," +
                        " Sum(BC_Util_Logs_biomass_wt) AS BC_Util_Logs_biomass_wt," +
                        " Sum(BC_Util_Chips_count) AS BC_Util_Chips_count," +
                        " Sum(BC_Util_Chips_TPA) AS BC_Util_Chips_TPA," +
                        " Sum(BC_Util_Chips_merch_wt) AS BC_Util_Chips_merch_wt," +
                        " Sum(BC_Util_Chips_merch_vol) AS BC_Util_Chips_merch_vol," +
                        " Sum(BC_Util_Chips_biomass_wt) AS BC_Util_Chips_biomass_wt," +
                        " Sum(BC_Util_Chips_biomass_vol) AS BC_Util_Chips_biomass_vol," +
                        " Sum(BC_NonUtil_Logs_count) AS BC_NonUtil_Logs_count," +
                        " Sum(BC_NonUtil_Logs_TPA) AS BC_NonUtil_Logs_TPA," +
                        " Sum(BC_NonUtil_Logs_merch_wt) AS BC_NonUtil_Logs_merch_wt," +
                        " Sum(BC_NonUtil_Logs_merch_vol) AS BC_NonUtil_Logs_merch_vol," +
                        " Sum(BC_NonUtil_Logs_biomass_vol) AS BC_NonUtil_Logs_biomass_vol," +
                        " Sum(BC_NonUtil_Logs_biomass_wt) AS BC_NonUtil_Logs_biomass_wt," +
                        " Sum(BC_NonUtil_Chips_count) AS BC_NonUtil_Chips_count," +
                        " Sum(BC_NonUtil_Chips_TPA) AS BC_NonUtil_Chips_TPA," +
                        " Sum(BC_NonUtil_Chips_merch_wt) AS BC_NonUtil_Chips_merch_wt," +
                        " Sum(BC_NonUtil_Chips_merch_vol) AS BC_NonUtil_Chips_merch_vol," +
                        " Sum(BC_NonUtil_Chips_biomass_wt) AS BC_NonUtil_Chips_biomass_wt," +
                        " Sum(BC_NonUtil_Chips_biomass_vol) AS BC_NonUtil_Chips_biomass_vol," +
                        " Sum(CHIPS_Util_Logs_count) AS CHIPS_Util_Logs_count," +
                        " Sum(CHIPS_Util_Logs_TPA) AS CHIPS_Util_Logs_TPA," +
                        " Sum(CHIPS_Util_Logs_merch_wt) AS CHIPS_Util_Logs_merch_wt," +
                        " Sum(CHIPS_Util_Logs_merch_vol) AS CHIPS_Util_Logs_merch_vol," +
                        " Sum(CHIPS_Util_Chips_count) AS CHIPS_Util_Chips_count," +
                        " Sum(CHIPS_Util_Chips_TPA) AS CHIPS_Util_Chips_TPA," +
                        " Sum(CHIPS_Util_Chips_merch_wt) AS CHIPS_Util_Chips_merch_wt," +
                        " Sum(CHIPS_Util_Chips_merch_vol) AS CHIPS_Util_Chips_merch_vol," +
                        " Sum(CHIPS_Util_Logs_biomass_wt) AS CHIPS_Util_Logs_biomass_wt," +
                        " Sum(CHIPS_Util_Logs_biomass_vol) AS CHIPS_Util_Logs_biomass_vol," +
                        " Sum(CHIPS_Util_Chips_biomass_wt) AS CHIPS_Util_Chips_biomass_wt," +
                        " Sum(CHIPS_Util_Chips_biomass_vol) AS CHIPS_Util_Chips_biomass_vol," +
                        " Sum(CHIPS_NonUtil_Logs_count) AS CHIPS_NonUtil_Logs_count," +
                        " Sum(CHIPS_NonUtil_Logs_TPA) AS CHIPS_NonUtil_Logs_TPA," +
                        " Sum(CHIPS_NonUtil_Logs_merch_wt) AS CHIPS_NonUtil_Logs_merch_wt," +
                        " Sum(CHIPS_NonUtil_Logs_merch_vol) AS CHIPS_NonUtil_Logs_merch_vol," +
                        " Sum(CHIPS_NonUtil_Chips_count) AS CHIPS_NonUtil_Chips_count," +
                        " Sum(CHIPS_NonUtil_Chips_TPA) AS CHIPS_NonUtil_Chips_TPA," +
                        " Sum(CHIPS_NonUtil_Chips_merch_wt) AS CHIPS_NonUtil_Chips_merch_wt," +
                        " Sum(CHIPS_NonUtil_Chips_merch_vol) AS CHIPS_NonUtil_Chips_merch_vol," +
                        " Sum(CHIPS_NonUtil_Logs_biomass_wt) AS CHIPS_NonUtil_Logs_biomass_wt," +
                        " Sum(CHIPS_NonUtil_Logs_biomass_vol) AS CHIPS_NonUtil_Logs_biomass_vol," +
                        " Sum(CHIPS_NonUtil_Chips_biomass_wt) AS CHIPS_NonUtil_Chips_biomass_wt," +
                        " Sum(CHIPS_NonUtil_Chips_biomass_vol) AS CHIPS_NonUtil_Chips_biomass_vol," +
                        " Sum(SMLOGS_Util_Logs_count) AS SMLOGS_Util_Logs_count," +
                        " Sum(SMLOGS_Util_Logs_TPA) AS SMLOGS_Util_Logs_TPA," +
                        " Sum(SMLOGS_Util_Logs_merch_wt) AS SMLOGS_Util_Logs_merch_wt," +
                        " Sum(SMLOGS_Util_Logs_merch_vol) AS SMLOGS_Util_Logs_merch_vol," +
                        " Sum(SMLOGS_Util_Logs_biomass_wt) AS SMLOGS_Util_Logs_biomass_wt," +
                        " Sum(SMLOGS_Util_Logs_biomass_vol) AS SMLOGS_Util_Logs_biomass_vol," +
                        " Sum(SMLOGS_Util_Chips_count) AS SMLOGS_Util_Chips_count," +
                        " Sum(SMLOGS_Util_Chips_TPA) AS SMLOGS_Util_Chips_TPA," +
                        " Sum(SMLOGS_Util_Chips_merch_wt) AS SMLOGS_Util_Chips_merch_wt," +
                        " Sum(SMLOGS_Util_Chips_merch_vol) AS SMLOGS_Util_Chips_merch_vol," +
                        " Sum(SMLOGS_Util_Chips_biomass_wt) AS SMLOGS_Util_Chips_biomass_wt," +
                        " Sum(SMLOGS_Util_Chips_biomass_vol) AS SMLOGS_Util_Chips_biomass_vol," +
						" Sum(SMLOGS_NonUtil_Logs_count) AS SMLOGS_NonUtil_Logs_count," +
                        " Sum(SMLOGS_NonUtil_Logs_TPA) AS SMLOGS_NonUtil_Logs_TPA," +
                        " Sum(SMLOGS_NonUtil_Logs_merch_wt) AS SMLOGS_NonUtil_Logs_merch_wt," +
                        " Sum(SMLOGS_NonUtil_Logs_merch_vol) AS SMLOGS_NonUtil_Logs_merch_vol," +
                        " Sum(SMLOGS_NonUtil_Logs_biomass_wt) AS SMLOGS_NonUtil_Logs_biomass_wt," +
                        " Sum(SMLOGS_NonUtil_Logs_biomass_vol) AS SMLOGS_NonUtil_Logs_biomass_vol," +
                        " Sum(SMLOGS_NonUtil_Chips_count) AS SMLOGS_NonUtil_Chips_count," +
                        " Sum(SMLOGS_NonUtil_Chips_TPA) AS SMLOGS_NonUtil_Chips_TPA," +
                        " Sum(SMLOGS_NonUtil_Chips_merch_wt) AS SMLOGS_NonUtil_Chips_merch_wt," +
                        " Sum(SMLOGS_NonUtil_Chips_merch_vol) AS SMLOGS_NonUtil_Chips_merch_vol," +
                        " Sum(SMLOGS_NonUtil_Chips_biomass_wt) AS SMLOGS_NonUtil_Chips_biomass_wt," +
                        " Sum(SMLOGS_NonUtil_Chips_biomass_vol) AS SMLOGS_NonUtil_Chips_biomass_vol," +
                        " Sum(LGLOGS_Util_Logs_count) AS LGLOGS_Util_Logs_count," +
                        " Sum(LGLOGS_Util_Logs_TPA) AS LGLOGS_Util_Logs_TPA," +
                        " Sum(LGLOGS_Util_Logs_merch_wt) AS LGLOGS_Util_Logs_merch_wt," +
                        " Sum(LGLOGS_Util_Logs_merch_vol) AS LGLOGS_Util_Logs_merch_vol," +
                        " Sum(LGLOGS_Util_Logs_biomass_wt) AS LGLOGS_Util_Logs_biomass_wt," +
                        " Sum(LGLOGS_Util_Logs_biomass_vol) AS LGLOGS_Util_Logs_biomass_vol," +
                        " Sum(LGLOGS_Util_Chips_count) AS LGLOGS_Util_Chips_count," +
                        " Sum(LGLOGS_Util_Chips_TPA) AS LGLOGS_Util_Chips_TPA," +
                        " Sum(LGLOGS_Util_Chips_merch_wt) AS LGLOGS_Util_Chips_merch_wt," +
                        " Sum(LGLOGS_Util_Chips_merch_vol) AS LGLOGS_Util_Chips_merch_vol, " +
                        " Sum(LGLOGS_Util_Chips_biomass_wt) AS LGLOGS_Util_Chips_biomass_wt," +
                        " Sum(LGLOGS_Util_Chips_biomass_vol) AS LGLOGS_Util_Chips_biomass_vol," +
                        " Sum(LGLOGS_NonUtil_Logs_count) AS LGLOGS_NonUtil_Logs_count," +
                        " Sum(LGLOGS_NonUtil_Logs_TPA) AS LGLOGS_NonUtil_Logs_TPA," +
                        " Sum(LGLOGS_NonUtil_Logs_merch_wt) AS LGLOGS_NonUtil_Logs_merch_wt," +
                        " Sum(LGLOGS_NonUtil_Logs_merch_vol) AS LGLOGS_NonUtil_Logs_merch_vol," +
                        " Sum(LGLOGS_NonUtil_Logs_biomass_wt) AS LGLOGS_NonUtil_Logs_biomass_wt," +
                        " Sum(LGLOGS_NonUtil_Logs_biomass_vol) AS LGLOGS_NonUtil_Logs_biomass_vol," +
                        " Sum(LGLOGS_NonUtil_Chips_count) AS LGLOGS_NonUtil_Chips_count," +
                        " Sum(LGLOGS_NonUtil_Chips_TPA) AS LGLOGS_NonUtil_Chips_TPA," +
                        " Sum(LGLOGS_NonUtil_Chips_merch_wt) AS LGLOGS_NonUtil_Chips_merch_wt," +
                        " Sum(LGLOGS_NonUtil_Chips_merch_vol) AS LGLOGS_NonUtil_Chips_merch_vol, " +
                        " Sum(LGLOGS_NonUtil_Chips_biomass_wt) AS LGLOGS_NonUtil_Chips_biomass_wt," +
                        " Sum(LGLOGS_NonUtil_Chips_biomass_vol) AS LGLOGS_NonUtil_Chips_biomass_vol " +
                        " INTO " + p_strIntoTableName + " " +
                        " FROM " + p_strFromTableName + " " +
                        " GROUP BY BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE";
                       
            }
            public static string SumHwdBinByPlotRx(string p_strFromTableName, string p_strIntoTableName)
            {
                return "SELECT BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE," +
                        " Sum(HWD_BC_Util_Logs_count) AS HWD_BC_Util_Logs_count," +
                        " Sum(HWD_BC_Util_Logs_TPA) AS HWD_BC_Util_Logs_TPA," +
                        " Sum(HWD_BC_Util_Logs_merch_wt) AS HWD_BC_Util_Logs_merch_wt," +
                        " Sum(HWD_BC_Util_Logs_merch_vol) AS HWD_BC_Util_Logs_merch_vol," +
                        " Sum(HWD_BC_Util_Logs_biomass_vol) AS HWD_BC_Util_Logs_biomass_vol," +
                        " Sum(HWD_BC_Util_Logs_biomass_wt) AS HWD_BC_Util_Logs_biomass_wt," +
                        " Sum(HWD_BC_Util_Chips_count) AS HWD_BC_Util_Chips_count," +
                        " Sum(HWD_BC_Util_Chips_TPA) AS HWD_BC_Util_Chips_TPA," +
                        " Sum(HWD_BC_Util_Chips_merch_wt) AS HWD_BC_Util_Chips_merch_wt," +
                        " Sum(HWD_BC_Util_Chips_merch_vol) AS HWD_BC_Util_Chips_merch_vol," +
                        " Sum(HWD_BC_Util_Chips_biomass_wt) AS HWD_BC_Util_Chips_biomass_wt," +
                        " Sum(HWD_BC_Util_Chips_biomass_vol) AS HWD_BC_Util_Chips_biomass_vol," +
                        " Sum(HWD_BC_NonUtil_Logs_count) AS HWD_BC_NonUtil_Logs_count," +
                        " Sum(HWD_BC_NonUtil_Logs_TPA) AS HWD_BC_NonUtil_Logs_TPA," +
                        " Sum(HWD_BC_NonUtil_Logs_merch_wt) AS HWD_BC_NonUtil_Logs_merch_wt," +
                        " Sum(HWD_BC_NonUtil_Logs_merch_vol) AS HWD_BC_NonUtil_Logs_merch_vol," +
                        " Sum(HWD_BC_NonUtil_Logs_biomass_vol) AS HWD_BC_NonUtil_Logs_biomass_vol," +
                        " Sum(HWD_BC_NonUtil_Logs_biomass_wt) AS HWD_BC_NonUtil_Logs_biomass_wt," +
                        " Sum(HWD_BC_NonUtil_Chips_count) AS HWD_BC_NonUtil_Chips_count," +
                        " Sum(HWD_BC_NonUtil_Chips_TPA) AS HWD_BC_NonUtil_Chips_TPA," +
                        " Sum(HWD_BC_NonUtil_Chips_merch_wt) AS HWD_BC_NonUtil_Chips_merch_wt," +
                        " Sum(HWD_BC_NonUtil_Chips_merch_vol) AS HWD_BC_NonUtil_Chips_merch_vol," +
                        " Sum(HWD_BC_NonUtil_Chips_biomass_wt) AS HWD_BC_NonUtil_Chips_biomass_wt," +
                        " Sum(HWD_BC_NonUtil_Chips_biomass_vol) AS HWD_BC_NonUtil_Chips_biomass_vol," +
                        " Sum(HWD_CHIPS_Util_Logs_count) AS HWD_CHIPS_Util_Logs_count," +
                        " Sum(HWD_CHIPS_Util_Logs_TPA) AS HWD_CHIPS_Util_Logs_TPA," +
                        " Sum(HWD_CHIPS_Util_Logs_merch_wt) AS HWD_CHIPS_Util_Logs_merch_wt," +
                        " Sum(HWD_CHIPS_Util_Logs_merch_vol) AS HWD_CHIPS_Util_Logs_merch_vol," +
                        " Sum(HWD_CHIPS_Util_Chips_count) AS HWD_CHIPS_Util_Chips_count," +
                        " Sum(HWD_CHIPS_Util_Chips_TPA) AS HWD_CHIPS_Util_Chips_TPA," +
                        " Sum(HWD_CHIPS_Util_Chips_merch_wt) AS HWD_CHIPS_Util_Chips_merch_wt," +
                        " Sum(HWD_CHIPS_Util_Chips_merch_vol) AS HWD_CHIPS_Util_Chips_merch_vol," +
                        " Sum(HWD_CHIPS_Util_Logs_biomass_wt) AS HWD_CHIPS_Util_Logs_biomass_wt," +
                        " Sum(HWD_CHIPS_Util_Logs_biomass_vol) AS HWD_CHIPS_Util_Logs_biomass_vol," +
                        " Sum(HWD_CHIPS_Util_Chips_biomass_wt) AS HWD_CHIPS_Util_Chips_biomass_wt," +
                        " Sum(HWD_CHIPS_Util_Chips_biomass_vol) AS HWD_CHIPS_Util_Chips_biomass_vol," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_count) AS HWD_CHIPS_NonUtil_Logs_count," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_TPA) AS HWD_CHIPS_NonUtil_Logs_TPA," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_merch_wt) AS HWD_CHIPS_NonUtil_Logs_merch_wt," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_merch_vol) AS HWD_CHIPS_NonUtil_Logs_merch_vol," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_count) AS HWD_CHIPS_NonUtil_Chips_count," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_TPA) AS HWD_CHIPS_NonUtil_Chips_TPA," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_merch_wt) AS HWD_CHIPS_NonUtil_Chips_merch_wt," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_merch_vol) AS HWD_CHIPS_NonUtil_Chips_merch_vol," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_biomass_wt) AS HWD_CHIPS_NonUtil_Logs_biomass_wt," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_biomass_vol) AS HWD_CHIPS_NonUtil_Logs_biomass_vol," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_biomass_wt) AS HWD_CHIPS_NonUtil_Chips_biomass_wt," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_biomass_vol) AS HWD_CHIPS_NonUtil_Chips_biomass_vol," +
                        " Sum(HWD_SMLOGS_Util_Logs_count) AS HWD_SMLOGS_Util_Logs_count," +
                        " Sum(HWD_SMLOGS_Util_Logs_TPA) AS HWD_SMLOGS_Util_Logs_TPA," +
                        " Sum(HWD_SMLOGS_Util_Logs_merch_wt) AS HWD_SMLOGS_Util_Logs_merch_wt," +
                        " Sum(HWD_SMLOGS_Util_Logs_merch_vol) AS HWD_SMLOGS_Util_Logs_merch_vol," +
                        " Sum(HWD_SMLOGS_Util_Logs_biomass_wt) AS HWD_SMLOGS_Util_Logs_biomass_wt," +
                        " Sum(HWD_SMLOGS_Util_Logs_biomass_vol) AS HWD_SMLOGS_Util_Logs_biomass_vol," +
                        " Sum(HWD_SMLOGS_Util_Chips_count) AS HWD_SMLOGS_Util_Chips_count," +
                        " Sum(HWD_SMLOGS_Util_Chips_TPA) AS HWD_SMLOGS_Util_Chips_TPA," +
                        " Sum(HWD_SMLOGS_Util_Chips_merch_wt) AS HWD_SMLOGS_Util_Chips_merch_wt," +
                        " Sum(HWD_SMLOGS_Util_Chips_merch_vol) AS HWD_SMLOGS_Util_Chips_merch_vol," +
                        " Sum(HWD_SMLOGS_Util_Chips_biomass_wt) AS HWD_SMLOGS_Util_Chips_biomass_wt," +
                        " Sum(HWD_SMLOGS_Util_Chips_biomass_vol) AS HWD_SMLOGS_Util_Chips_biomass_vol," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_count) AS HWD_SMLOGS_NonUtil_Logs_count," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_TPA) AS HWD_SMLOGS_NonUtil_Logs_TPA," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_merch_wt) AS HWD_SMLOGS_NonUtil_Logs_merch_wt," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_merch_vol) AS HWD_SMLOGS_NonUtil_Logs_merch_vol," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_biomass_wt) AS HWD_SMLOGS_NonUtil_Logs_biomass_wt," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_biomass_vol) AS HWD_SMLOGS_NonUtil_Logs_biomass_vol," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_count) AS HWD_SMLOGS_NonUtil_Chips_count," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_TPA) AS HWD_SMLOGS_NonUtil_Chips_TPA," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_merch_wt) AS HWD_SMLOGS_NonUtil_Chips_merch_wt," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_merch_vol) AS HWD_SMLOGS_NonUtil_Chips_merch_vol," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_biomass_wt) AS HWD_SMLOGS_NonUtil_Chips_biomass_wt," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_biomass_vol) AS HWD_SMLOGS_NonUtil_Chips_biomass_vol," +
                        " Sum(HWD_LGLOGS_Util_Logs_count) AS HWD_LGLOGS_Util_Logs_count," +
                        " Sum(HWD_LGLOGS_Util_Logs_TPA) AS HWD_LGLOGS_Util_Logs_TPA," +
                        " Sum(HWD_LGLOGS_Util_Logs_merch_wt) AS HWD_LGLOGS_Util_Logs_merch_wt," +
                        " Sum(HWD_LGLOGS_Util_Logs_merch_vol) AS HWD_LGLOGS_Util_Logs_merch_vol," +
                        " Sum(HWD_LGLOGS_Util_Logs_biomass_wt) AS HWD_LGLOGS_Util_Logs_biomass_wt," +
                        " Sum(HWD_LGLOGS_Util_Logs_biomass_vol) AS HWD_LGLOGS_Util_Logs_biomass_vol," +
                        " Sum(HWD_LGLOGS_Util_Chips_count) AS HWD_LGLOGS_Util_Chips_count," +
                        " Sum(HWD_LGLOGS_Util_Chips_TPA) AS HWD_LGLOGS_Util_Chips_TPA," +
                        " Sum(HWD_LGLOGS_Util_Chips_merch_wt) AS HWD_LGLOGS_Util_Chips_merch_wt," +
                        " Sum(HWD_LGLOGS_Util_Chips_merch_vol) AS HWD_LGLOGS_Util_Chips_merch_vol," +
                        " Sum(HWD_LGLOGS_Util_Chips_biomass_wt) AS HWD_LGLOGS_Util_Chips_biomass_wt," +
                        " Sum(HWD_LGLOGS_Util_Chips_biomass_vol) AS HWD_LGLOGS_Util_Chips_biomass_vol, " +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_count) AS HWD_LGLOGS_NonUtil_Logs_count," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_TPA) AS HWD_LGLOGS_NonUtil_Logs_TPA," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_merch_wt) AS HWD_LGLOGS_NonUtil_Logs_merch_wt," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_merch_vol) AS HWD_LGLOGS_NonUtil_Logs_merch_vol," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_biomass_wt) AS HWD_LGLOGS_NonUtil_Logs_biomass_wt," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_biomass_vol) AS HWD_LGLOGS_NonUtil_Logs_biomass_vol," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_count) AS HWD_LGLOGS_NonUtil_Chips_count," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_TPA) AS HWD_LGLOGS_NonUtil_Chips_TPA," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_merch_wt) AS HWD_LGLOGS_NonUtil_Chips_merch_wt," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_merch_vol) AS HWD_LGLOGS_NonUtil_Chips_merch_vol," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_biomass_wt) AS HWD_LGLOGS_NonUtil_Chips_biomass_wt," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_biomass_vol) AS HWD_LGLOGS_NonUtil_Chips_biomass_vol " +
                        " INTO " + p_strIntoTableName + " " +
                        " FROM " + p_strFromTableName + " " +
                        " GROUP BY BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE";
                       
            }
            public static string InitializeOutputTableValues(string p_strOutputTableName)
            {
               
                return
                    "UPDATE " + p_strOutputTableName + " o " +
                    "SET o.[CHIPS TPA] = 0," + 
                        "o.[CHIPS Average Vol (ft3)] = 0," + 
                        "o.[CHIPS Average Weight (tons)] = 0," +
                        "o.[CHIPS Average Density (lbs/ft3)] = 0," + 
                        "o.[CHIPS Hwd Proportion] = 0," + 
                        "o.[CHIPS Chip Fraction] = 0," +
                        "o.[CHIPS utilized logs (ft3)] = 0," +
                        "o.[CHIPS utilized chips (tons)] = 0," +
                        "o.[SMLOGS TPA] = 0," +
                        "o.[SMLOGS Average Vol (ft3)] = 0," +
                        "o.[SMLOGS Average Weight (tons)] = 0," +
                        "o.[SMLOGS Average Density (lbs/ft3)] = 0," +
                        "o.[SMLOGS Hwd Proportion] = 0," +
                        "o.[SMLOGS Chip Fraction] = 0," +
                        "o.[SMLOGS utilized logs (ft3)] = 0," +
                        "o.[SMLOGS utilized chips (tons)] = 0," +
                        "o.[LGLOGS TPA] = 0," +
                        "o.[LGLOGS Average Vol (ft3)] = 0," +
                        "o.[LGLOGS Average Weight (tons)] = 0," +
                        "o.[LGLOGS Average Density (lbs/ft3)] =0," +
                        "o.[LGLOGS Hwd Proportion] =0," +
                        "o.[LGLOGS Chip Fraction] = 0," +
                        "o.[LGLOGS utilized logs (ft3)] = 0," +
                        "o.[LGLOGS utilized chips (tons)] = 0," +
                        "o.[TOTAL TPA] = 0," +
                        "o.[TOTAL Average Vol (ft3)] = 0," +
                        "o.[TOTAL Average Weight (tons)] = 0," +
                        "o.[TOTAL Average Density (lbs/ft3)] = 0," +
                        "o.[TOTAL Hwd Proportion] = 0," +
                        "o.[TOTAL Chip Fraction] = 0," +
                        "o.[TOTAL utilized logs (ft3)] = 0," +
                        "o.[TOTAL utilized chips (tons)] = 0," +
                        "o.[BRUSH CUT utilized logs (ft3)] = 0," +
                        "o.[BRUSH CUT utilized chips (tons)] = 0," +
                        "o.[BRUSH CUT not utilized TPA] = 0," +
                        "o.[BRUSH CUT not utilized Average Vol (ft3)] = 0," +
                        "o.[BRUSH CUT not utilized Average Weight (tons)] = 0, " +
                        "o.[BRUSH CUT not utilized Average Density (lbs/ft3)] = 0, " +
                        "o.[BRUSH CUT not utilized Hwd Proportion] = 0," +
                        "o.[BRUSH CUT not utilized Chip Fraction] = 0," +
                        "o.[BRUSH CUT not utilized logs (ft3)] = 0," +
                        "o.[BRUSH CUT not utilized chips (tons)] = 0," + 
                        "o.gis_yard_dist=1," + 
                        "o.slope=0," + 
                        "o.elev=0";

           
            }
            public static string UpdateOutputTableFromBinTables(string p_strOutputTableName, string p_strBinTotalsTableName, string p_strHwdBinTotalsTableName)
            {
                return
                    "UPDATE (" + p_strOutputTableName + " o " +
                    "INNER JOIN " + p_strBinTotalsTableName + " b " +
                    "ON (o.biosum_cond_id = b.biosum_cond_id AND " +
                        "o.rxpackage = b.rxpackage AND " +
                        "o.rx = b.rx AND " +
                        "o.rxcycle = b.rxcycle)) " +
                    "INNER JOIN " + p_strHwdBinTotalsTableName + " h " +
                    "ON (b.biosum_cond_id = h.biosum_cond_id AND " +
                        "b.rxpackage = h.rxpackage AND " +
                        "b.rx = h.rx AND " +
                        "b.rxcycle = h.rxcycle) " +
                    "SET o.[CHIPS TPA] = b.CHIPS_Util_Chips_TPA + b.SMLOGS_Util_Chips_TPA + b.LGLOGS_Util_Chips_TPA," +
                        "o.[CHIPS Average Vol (ft3)] = " +
                          "IIF(b.CHIPS_Util_Chips_TPA+ b.SMLOGS_Util_Chips_TPA + b.LGLOGS_Util_Chips_TPA > 0," +
                          "(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Chips_biomass_vol+b.LGLOGS_Util_Chips_biomass_vol)/(b.CHIPS_Util_Chips_TPA+ b.SMLOGS_Util_Chips_TPA + b.LGLOGS_Util_Chips_TPA),0)," +
                        "o.[CHIPS Average Weight (tons)] = " +
                          "IIF(b.CHIPS_Util_Chips_TPA+ b.SMLOGS_Util_Chips_TPA + b.LGLOGS_Util_Chips_TPA>0," +
                          "(b.CHIPS_Util_Chips_biomass_wt+b.SMLOGS_Util_Chips_biomass_wt+b.LGLOGS_Util_Chips_biomass_wt)/(b.CHIPS_Util_Chips_TPA+ b.SMLOGS_Util_Chips_TPA + b.LGLOGS_Util_Chips_TPA),0)," +
                        "o.[CHIPS Average Density (lbs/ft3)] = IIF(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Chips_biomass_vol+b.LGLOGS_Util_Chips_biomass_vol=0,0,((b.CHIPS_Util_Chips_biomass_wt+b.SMLOGS_Util_Chips_biomass_wt+b.LGLOGS_Util_Chips_biomass_wt)*2000)/(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Chips_biomass_vol+b.LGLOGS_Util_Chips_biomass_vol))," +
                        "o.[CHIPS Hwd Proportion] = IIF(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Chips_biomass_vol+b.LGLOGS_Util_Chips_biomass_vol=0,0,(h.HWD_CHIPS_Util_Chips_biomass_vol+h.HWD_SMLOGS_Util_Chips_biomass_vol+h.HWD_LGLOGS_Util_Chips_biomass_vol)/(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Chips_biomass_vol+b.LGLOGS_Util_Chips_biomass_vol))," +
                         "o.[CHIPS Chip Fraction] = IIf(o.[CHIPS Average Vol (ft3)]=0,0,100)," +
                        "o.[CHIPS utilized logs (ft3)] = b.CHIPS_Util_Logs_biomass_vol+b.SMLOGS_Util_Chips_biomass_vol+b.LGLOGS_Util_Chips_biomass_vol," +
                        "o.[CHIPS utilized chips (tons)] = b.CHIPS_Util_Chips_biomass_wt+b.SMLOGS_Util_Chips_biomass_wt+b.LGLOGS_Util_Chips_biomass_wt," +
                        "o.[SMLOGS TPA] = b.SMLOGS_Util_Logs_TPA+b.SMLOGS_Util_Chips_TPA," +
                        "o.[SMLOGS Average Vol (ft3)] = IIF(b.SMLOGS_Util_Logs_TPA+b.SMLOGS_Util_Chips_TPA=0,0,(b.SMLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol)/(b.SMLOGS_Util_Logs_TPA+b.SMLOGS_Util_Chips_TPA))," +
                        "o.[SMLOGS Average Weight (tons)] = IIF(b.SMLOGS_Util_Logs_TPA+b.SMLOGS_Util_Chips_TPA=0,0,(b.SMLOGS_Util_Logs_merch_wt+b.SMLOGS_Util_Chips_merch_wt)/(b.SMLOGS_Util_Logs_TPA+b.SMLOGS_Util_Chips_TPA))," +
                        "o.[SMLOGS Average Density (lbs/ft3)] = IIF(b.SMLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol=0,0,((b.SMLOGS_Util_Logs_merch_wt+b.SMLOGS_Util_Chips_merch_wt)*2000)/(b.SMLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol))," +
                        "o.[SMLOGS Hwd Proportion] = IIF(b.SMLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol=0,0,(h.HWD_SMLOGS_Util_Logs_merch_vol+h.HWD_SMLOGS_Util_Chips_merch_vol)/(b.SMLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol))," +
                        "o.[SMLOGS Chip Fraction] = 0," +
                        "o.[SMLOGS utilized logs (ft3)] = b.SMLOGS_Util_Logs_merch_vol," +
                        "o.[SMLOGS utilized chips (tons)] = b.SMLOGS_Util_Chips_biomass_wt + b.SMLOGS_Util_Chips_merch_wt," +
                        "o.[LGLOGS TPA] = b.LGLOGS_Util_Logs_TPA + b.LGLOGS_Util_Chips_TPA," +
                        "o.[LGLOGS Average Vol (ft3)] = IIF(b.LGLOGS_Util_Logs_TPA+b.LGLOGS_Util_Chips_TPA=0,0,(b.LGLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Chips_merch_vol)/(b.LGLOGS_Util_Logs_TPA+b.LGLOGS_Util_Chips_TPA))," +
                        "o.[LGLOGS Average Weight (tons)] = IIF(b.LGLOGS_Util_Logs_TPA+b.LGLOGS_Util_Chips_TPA=0,0,(b.LGLOGS_Util_Logs_merch_wt+b.LGLOGS_Util_Chips_merch_wt)/(b.LGLOGS_Util_Logs_TPA+b.LGLOGS_Util_Chips_TPA))," +
                        "o.[LGLOGS Average Density (lbs/ft3)] =IIF(b.LGLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Chips_merch_vol=0,0,((b.LGLOGS_Util_Logs_merch_wt+b.LGLOGS_Util_Chips_merch_wt)*2000)/(b.LGLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Chips_merch_vol))," +
                        "o.[LGLOGS Hwd Proportion] =IIF(b.LGLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Chips_merch_vol=0,0,(h.HWD_LGLOGS_Util_Logs_merch_vol+h.HWD_LGLOGS_Util_Chips_merch_vol)/(b.LGLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Chips_merch_vol))," +
                        "o.[LGLOGS Chip Fraction] = 0," +
                        "o.[LGLOGS utilized logs (ft3)] = b.LGLOGS_Util_Logs_merch_vol," +
                        "o.[LGLOGS utilized chips (tons)] = b.LGLOGS_Util_Chips_biomass_wt + b.LGLOGS_Util_Chips_merch_wt," +
                        "o.[TOTAL TPA] = b.CHIPS_Util_Chips_TPA+b.SMLOGS_Util_Logs_TPA+b.LGLOGS_Util_Logs_TPA+b.SMLOGS_Util_Chips_TPA+b.LGLOGS_Util_Chips_TPA," +
                        "o.[TOTAL Average Vol (ft3)] = IIF(b.CHIPS_Util_Chips_TPA+b.SMLOGS_Util_Logs_TPA+b.LGLOGS_Util_Logs_TPA+b.SMLOGS_Util_Chips_TPA+b.LGLOGS_Util_Chips_TPA=0,0,(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol+b.LGLOGS_Util_Chips_merch_vol)/" +
                                "(b.CHIPS_Util_Chips_TPA+b.SMLOGS_Util_Logs_TPA+b.LGLOGS_Util_Logs_TPA+b.SMLOGS_Util_Chips_TPA+b.LGLOGS_Util_Chips_TPA))," +
                        "o.[TOTAL Average Weight (tons)] = IIF(b.CHIPS_Util_Chips_TPA+b.SMLOGS_Util_Logs_TPA+b.LGLOGS_Util_Logs_TPA=0,0,(b.CHIPS_Util_Chips_biomass_wt+b.SMLOGS_Util_Logs_merch_wt+b.LGLOGS_Util_Logs_merch_wt)/" +
                                "(b.CHIPS_Util_Chips_TPA+b.SMLOGS_Util_Logs_TPA+b.LGLOGS_Util_Logs_TPA))," +
                        "o.[TOTAL Average Density (lbs/ft3)] = IIF(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol+b.LGLOGS_Util_Chips_merch_vol=0,0,((b.CHIPS_Util_Chips_biomass_wt+b.SMLOGS_Util_Logs_merch_wt+b.LGLOGS_Util_Logs_merch_wt+b.SMLOGS_Util_Chips_merch_wt+b.LGLOGS_Util_Chips_merch_wt)*2000)/" +
                                "(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol+b.LGLOGS_Util_Chips_merch_vol))," +
                        "o.[TOTAL Hwd Proportion] = IIF(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol+b.LGLOGS_Util_Chips_merch_vol=0,0,(h.HWD_CHIPS_Util_Chips_biomass_vol+h.HWD_SMLOGS_Util_Logs_merch_vol+h.HWD_LGLOGS_Util_Logs_merch_vol+h.HWD_SMLOGS_Util_Chips_merch_vol+h.HWD_LGLOGS_Util_Chips_merch_vol)/" +
                                "(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol+b.LGLOGS_Util_Chips_merch_vol))," +
                        "o.[TOTAL Chip Fraction] = 0," +
                        "o.[TOTAL utilized logs (ft3)] = b.SMLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Logs_merch_vol," +
                        "o.[TOTAL utilized chips (tons)] = b.CHIPS_Util_Chips_biomass_wt+b.SMLOGS_Util_Chips_biomass_wt+b.LGLOGS_Util_Chips_biomass_wt+b.SMLOGS_Util_Chips_merch_wt+b.LGLOGS_Util_Chips_merch_wt," +
                        "o.[BRUSH CUT utilized logs (ft3)] = b.BC_Util_Logs_biomass_vol," +
                        "o.[BRUSH CUT utilized chips (tons)] = b.BC_Util_Chips_biomass_wt," +
                        "o.[BRUSH CUT not utilized TPA] = b.BC_NonUtil_Chips_TPA + b.BC_NonUtil_Logs_TPA," +
                        "o.[BRUSH CUT not utilized Average Vol (ft3)] = " +
                                "IIF(b.BC_NonUtil_Chips_TPA+ b.BC_NonUtil_Logs_TPA > 0," +
                                    "(b.BC_NonUtil_Chips_biomass_vol)/(b.BC_NonUtil_Chips_TPA+ b.BC_NonUtil_Logs_TPA),0)," +
                        "o.[BRUSH CUT not utilized Average Weight (tons)] = " +
                                "IIF(b.BC_NonUtil_Chips_TPA+ b.BC_NonUtil_Logs_TPA>0," +
                                   "(b.BC_NonUtil_Chips_biomass_wt+b.BC_NonUtil_Logs_biomass_wt)/(b.BC_NonUtil_Chips_TPA + b.BC_NonUtil_Logs_TPA),0)," +
                        "o.[BRUSH CUT not utilized Average Density (lbs/ft3)] = " +
                                "IIF(b.BC_NonUtil_Chips_biomass_vol+b.BC_NonUtil_Logs_biomass_vol=0,0,((b.BC_NonUtil_Chips_biomass_wt+b.BC_NonUtil_Logs_biomass_wt)*2000)/(b.BC_NonUtil_Chips_biomass_vol+b.BC_NonUtil_Logs_biomass_vol))," +
                        "o.[BRUSH CUT not utilized Hwd Proportion] = " +
                                "IIF(b.BC_NonUtil_Chips_biomass_vol+b.BC_NonUtil_Logs_biomass_vol=0,0,(h.HWD_BC_NonUtil_Chips_biomass_vol+h.HWD_BC_NonUtil_Logs_biomass_vol)/(b.BC_NonUtil_Chips_biomass_vol+b.BC_NonUtil_Logs_biomass_vol))," +
                        "o.[BRUSH CUT not utilized Chip Fraction] = IIf(o.[BRUSH CUT not utilized Average Vol (ft3)]=0,0,100)," +
                        "o.[BRUSH CUT not utilized logs (ft3)] = b.BC_NonUtil_Logs_biomass_vol," +
                        "o.[BRUSH CUT not utilized chips (tons)] = b.BC_NonUtil_Chips_biomass_wt";

            }
            public static string SumBinSumTableByPlotRxSpcGrpDbhGrp(string p_strIntoTableName, string p_strBinSumTableName, string p_strHwdBinSumTableName)
            {
                return
                    "SELECT b.BIOSUM_COND_ID,b.RXPACKAGE, b.RX, b.RXCYCLE, " +
                           "b.species_group, b.diam_group," +
                           "SUM(b.[BC_Util_Logs_merch_wt] + b.[CHIPS_Util_Logs_merch_wt] " +
                           "+ b.[SMLOGS_Util_Logs_merch_wt] + b.[LGLOGS_Util_Logs_merch_wt] " +
                           ") " +
                           "AS MERCH_WT_GT," +
                           "SUM(b.[BC_Util_Logs_merch_vol] + b.[CHIPS_Util_Logs_merch_vol] " +
                           "+ b.[SMLOGS_Util_Logs_merch_vol] + b.[LGLOGS_Util_Logs_merch_vol] " +
                           ") " +
                           "AS MERCH_VOL_CF," +
                           "SUM(b.[BC_Util_Logs_biomass_wt] + b.[CHIPS_Util_Logs_biomass_wt] " +
                           "+ b.[SMLOGS_Util_Logs_biomass_wt] + b.[LGLOGS_Util_Logs_biomass_wt] " +
                           "+ b.[SMLOGS_Util_Chips_biomass_wt] + b.[LGLOGS_Util_Chips_biomass_wt] " +
                           "+ b.[CHIPS_Util_Chips_biomass_wt]+ b.[SMLOGS_Util_Chips_merch_wt]) " +
                           "AS CHIP_WT_GT," +
                           "SUM(b.[BC_Util_Logs_biomass_vol] + b.[CHIPS_Util_Logs_biomass_vol] " +
                           "+ b.[SMLOGS_Util_Logs_biomass_vol] + b.[LGLOGS_Util_Logs_biomass_vol] " +
                           "+ b.[SMLOGS_Util_Chips_biomass_vol] + b.[LGLOGS_Util_Chips_biomass_vol] " +
                           "+ b.[CHIPS_Util_Chips_biomass_vol] + b.[LGLOGS_Util_Chips_merch_vol]) " +
                           "AS CHIP_VOL_CF," +
                           "SUM(b.[BC_NonUtil_Logs_merch_wt] + b.[CHIPS_NonUtil_Logs_merch_wt] " +
                           "+ b.[SMLOGS_NonUtil_Logs_merch_wt] + b.[LGLOGS_NonUtil_Logs_merch_wt] " +
                           ") " +
                           "AS NOT_UTILIZED_MERCH_WT_GT," +
                           "SUM(b.[BC_NonUtil_Logs_merch_vol] + b.[CHIPS_NonUtil_Logs_merch_vol] " +
                           "+ b.[SMLOGS_NonUtil_Logs_merch_vol] + b.[LGLOGS_NonUtil_Logs_merch_vol] " +
                           ") " +
                           "AS NOT_UTILIZED_MERCH_VOL_CF," +
                           "SUM(b.[BC_NonUtil_Logs_biomass_wt] + b.[CHIPS_NonUtil_Logs_biomass_wt] " +
                           "+ b.[SMLOGS_NonUtil_Logs_biomass_wt] + b.[LGLOGS_NonUtil_Logs_biomass_wt] " +
                           "+ b.[SMLOGS_NonUtil_Chips_biomass_wt] + b.[LGLOGS_NonUtil_Chips_biomass_wt] " +
                           "+ b.[SMLOGS_NonUtil_Chips_merch_wt]) " +
                           "AS NOT_UTILIZED_CHIP_WT_GT," +
                           "SUM(b.[BC_NonUtil_Logs_biomass_vol] + b.[CHIPS_NonUtil_Logs_biomass_vol] " +
                           "+ b.[SMLOGS_NonUtil_Logs_biomass_vol] + b.[LGLOGS_NonUtil_Logs_biomass_vol] " +
                           "+ b.[SMLOGS_NonUtil_Chips_biomass_vol] + b.[LGLOGS_NonUtil_Chips_biomass_vol] " +
                           "+ b.[LGLOGS_NonUtil_Chips_merch_vol]) " +
                           "AS NOT_UTILIZED_CHIP_VOL_CF, " +
                           "SUM(b.[BC_NonUtil_Chips_merch_wt] + b.[BC_NonUtil_Chips_biomass_wt]) " +
                           "AS BC_WT_GT, " +
                           "SUM(b.[BC_NonUtil_Chips_merch_vol] + b.[BC_NonUtil_Chips_biomass_vol]) " +
                           "AS BC_VOL_CF " +
                           "INTO " + p_strIntoTableName + " " +
                           "FROM " + p_strBinSumTableName + " b " +
                           "WHERE (((b.diam_group) Is Not Null)) " +
                           "GROUP BY b.BIOSUM_COND_ID, b.rxpackage, b.RX, b.rxcycle, b.species_group, b.diam_group";

            }
           
            public static string AppendToSumSoftwoodBinSumTableByPlotRxSpcGrpDbhGrp(string p_strTable, string p_strSoftwoodBinSumTableName,string p_strHardwoodBinSumTableName)
            {
                return "INSERT INTO " + p_strTable + " " +
                       "SELECT b.BIOSUM_COND_ID,b.RXPACKAGE, b.RX, b.RXCYCLE, b.species_group," +
                              "b.diam_group," +
                               "SUM(b.[BC_Util_Logs_merch_wt] + " +
                                   "b.[CHIPS_Util_Logs_merch_wt] + " +
                                   "b.[SMLOGS_Util_Logs_merch_wt] + " +
                                   "b.[LGLOGS_Util_Logs_merch_wt]) AS MERCH_WT_GT," +
                               "SUM(b.[BC_Util_Logs_merch_vol] + " +
                                   "b.[CHIPS_Util_Logs_merch_vol] + " +
                                   "b.[SMLOGS_Util_Logs_merch_vol] + " +
                                   "b.[LGLOGS_Util_Logs_merch_vol]) AS MERCH_VOL_CF," +
                               "SUM(b.[BC_Util_Logs_biomass_wt] + " +
                                   "b.[CHIPS_Util_Logs_biomass_wt] + " +
                                   "b.[SMLOGS_Util_Logs_biomass_wt] + " +
                                   "b.[LGLOGS_Util_Logs_biomass_wt] + " +
                                   "b.[SMLOGS_Util_Chips_biomass_wt] + " +
                                   "b.[LGLOGS_Util_Chips_biomass_wt] + " +
                                   "b.[SMLOGS_Util_Chips_merch_wt]) AS CHIP_WT_GT," +
                               "SUM(b.[BC_Util_Logs_biomass_vol] + b.[CHIPS_Util_Logs_biomass_vol] + " +
                                   "b.[SMLOGS_Util_Logs_biomass_vol] + " +
                                   "b.[LGLOGS_Util_Logs_biomass_vol] + " +
                                   "b.[SMLOGS_Util_Chips_biomass_vol] + " +
                                   "b.[LGLOGS_Util_Chips_biomass_vol]  + " +
                                   "b.[LGLOGS_Util_Chips_merch_vol]) AS CHIP_VOL_CF " +
                                "SUM(b.[BC_NonUtil_Logs_merch_wt] + " +
                                   "b.[CHIPS_NonUtil_Logs_merch_wt] + " +
                                   "b.[SMLOGS_NonUtil_Logs_merch_wt] + " +
                                   "b.[LGLOGS_NonUtil_Logs_merch_wt]) AS NOT_UTILIZED_MERCH_WT_GT," +
                                "SUM(b.[BC_NonUtil_Logs_merch_vol] + " +
                                   "b.[CHIPS_NonUtil_Logs_merch_vol] + " +
                                   "b.[SMLOGS_NonUtil_Logs_merch_vol] + " +
                                   "b.[LGLOGS_NonUtil_Logs_merch_vol]) AS NOT_UTILIZED_MERCH_VOL_CF," +
                                "SUM(b.[BC_NonUtil_Logs_biomass_wt] + " +
                                   "b.[CHIPS_NonUtil_Logs_biomass_wt] + " +
                                   "b.[SMLOGS_NonUtil_Logs_biomass_wt] + " +
                                   "b.[LGLOGS_NonUtil_Logs_biomass_wt] + " +
                                   "b.[SMLOGS_NonUtil_Chips_biomass_wt] + " +
                                   "b.[LGLOGS_NonUtil_Chips_biomass_wt] + " +
                                   "b.[SMLOGS_NonUtil_Chips_merch_wt]) AS NOT_UTILIZED_CHIP_WT_GT," +
                                "SUM(b.[BC_NonUtil_Logs_biomass_vol] + b.[CHIPS_NonUtil_Logs_biomass_vol] + " +
                                   "b.[SMLOGS_NonUtil_Logs_biomass_vol] + " +
                                   "b.[LGLOGS_NonUtil_Logs_biomass_vol] + " +
                                   "b.[SMLOGS_NonUtil_Chips_biomass_vol] + " +
                                   "b.[LGLOGS_NonUtil_Chips_biomass_vol]  + " +
                                   "b.[LGLOGS_NonUtil_Chips_merch_vol]) AS NOT_UTILIZED_CHIP_VOL_CF " +
                       "FROM " + p_strSoftwoodBinSumTableName + " b " +
                       "GROUP BY b.BIOSUM_COND_ID, b.rxpackage, b.RX, b.rxcycle," +
                                "b.species_group, b.diam_group";
                     
            }
            public static string AppendToSumHardwoodBinSumTableByPlotRxSpcGrpDbhGrp(string p_strTable, string p_strSoftwoodBinSumTableName, string p_strHardwoodBinSumTableName)
            {
                return "INSERT INTO " + p_strTable + " " +
                       "SELECT h.BIOSUM_COND_ID,h.RXPACKAGE, h.RX, h.RXCYCLE, h.species_group," +
                              "h.diam_group," +
                              "SUM(h.[HWD_BC_Util_Logs_merch_wt] + " +
                                  "h.[HWD_CHIPS_Util_Logs_merch_wt] + " +
                                  "h.[HWD_SMLOGS_Util_Logs_merch_wt] + " +
                                  "h.[HWD_LGLOGS_Util_Logs_merch_wt]) AS MERCH_WT_GT," +
                              "SUM(h.[HWD_BC_Util_Logs_merch_vol] + " +
                                  "h.[HWD_CHIPS_Util_Logs_merch_vol] + " +
                                  "h.[HWD_SMLOGS_Util_Logs_merch_vol] + " +
                                  "h.[HWD_LGLOGS_Util_Logs_merch_vol]) AS MERCH_VOL_CF," +
                             "SUM(h.[HWD_BC_Util_Logs_biomass_wt] + " +
                                 "h.[HWD_CHIPS_Util_Logs_biomass_wt] + " +
                                 "h.[HWD_SMLOGS_Util_Logs_biomass_wt] + " +
                                 "h.[HWD_LGLOGS_Util_Logs_biomass_wt] + " +
                                 "h.[HWD_SMLOGS_Util_Chips_biomass_wt] + " +
                                 "h.[HWD_LGLOGS_Util_Chips_biomass_wt] + " +
                                 "h.[HWD_SMLOGS_Util_Chips_merch_wt]) AS CHIP_WT_GT," +
                             "SUM(h.[HWD_BC_Util_Logs_biomass_vol] + " +
                                 "h.[HWD_CHIPS_Util_Logs_biomass_vol] + " +
                                 "h.[HWD_SMLOGS_Util_Logs_biomass_vol] + " +
                                 "h.[HWD_LGLOGS_Util_Logs_biomass_vol] + " +
                                 "h.[HWD_SMLOGS_Util_Chips_biomass_vol] + " +
                                 "h.[HWD_LGLOGS_Util_Chips_biomass_vol] +  " +
                                 "h.[HWD_LGLOGS_Util_Chips_merch_vol]) AS CHIP_VOL_CF," +
                             "SUM(h.[HWD_BC_NonUtil_Logs_merch_wt] + " +
                                 "h.[HWD_CHIPS_NonUtil_Logs_merch_wt] + " +
                                 "h.[HWD_SMLOGS_NonUtil_Logs_merch_wt] + " +
                                 "h.[HWD_LGLOGS_NonUtil_Logs_merch_wt]) AS NOT_UTILIZED_MERCH_WT_GT," +
                             "SUM(h.[HWD_BC_NonUtil_Logs_merch_vol] + " +
                                  "h.[HWD_CHIPS_NonUtil_Logs_merch_vol] + " +
                                  "h.[HWD_SMLOGS_NonUtil_Logs_merch_vol] + " +
                                  "h.[HWD_LGLOGS_NonUtil_Logs_merch_vol]) AS NOT_UTILIZED_MERCH_VOL_CF," +
                             "SUM(h.[HWD_BC_NonUtil_Logs_biomass_wt] + " +
                                 "h.[HWD_CHIPS_NonUtil_Logs_biomass_wt] + " +
                                 "h.[HWD_SMLOGS_NonUtil_Logs_biomass_wt] + " +
                                 "h.[HWD_LGLOGS_NonUtil_Logs_biomass_wt] + " +
                                 "h.[HWD_SMLOGS_NonUtil_Chips_biomass_wt] + " +
                                 "h.[HWD_LGLOGS_NonUtil_Chips_biomass_wt] + " +
                                 "h.[HWD_SMLOGS_NonUtil_Chips_merch_wt]) AS NOT_UTILIZED_CHIP_WT_GT," +
                             "SUM(h.[HWD_BC_NonUtil_Logs_biomass_vol] + " +
                                 "h.[HWD_CHIPS_NonUtil_Logs_biomass_vol] + " +
                                 "h.[HWD_SMLOGS_NonUtil_Logs_biomass_vol] + " +
                                 "h.[HWD_LGLOGS_NonUtil_Logs_biomass_vol] + " +
                                 "h.[HWD_SMLOGS_NonUtil_Chips_biomass_vol] + " +
                                 "h.[HWD_LGLOGS_NonUtil_Chips_biomass_vol] +  " +
                                 "h.[HWD_LGLOGS_NonUtil_Chips_merch_vol]) AS NOT_UTILIZED_CHIP_VOL_CF," +
                       "FROM " + p_strHardwoodBinSumTableName + " h " +
                        "GROUP BY h.BIOSUM_COND_ID, h.rxpackage, h.RX, h.rxcycle," +
                                "h.species_group, h.diam_group";
                       
            }
            public static string AppendToTreeVolVal(string p_strSourceTable, string p_strDestTable,string p_strDateTimeCreated)
            {
                return
                    "INSERT INTO " + p_strDestTable + " " +
                        "(biosum_cond_id,rxpackage,rx,rxcycle,species_group,diam_group, merch_wt_gt," +
                         "merch_vol_cf, chip_wt_gt, chip_vol_cf, bc_wt_gt, bc_vol_cf, DateTimeCreated) " +
                         "SELECT DISTINCT t.BIOSUM_COND_ID, t.rxpackage,t.rx,t.rxcycle," +
                                        "t.species_group,t.diam_group," +
                                        "t.MERCH_WT_GT,t.MERCH_VOL_CF," +
                                        "t.CHIP_WT_GT,t.CHIP_VOL_CF," +
                                        "t.BC_WT_GT,t.BC_VOL_CF," +
                                        "'" + p_strDateTimeCreated + "' " + 
                         "FROM " + p_strSourceTable + " t " +
                         "WHERE (((t.MERCH_WT_GT)<>0) OR ((t.MERCH_VOL_CF)<>0) OR " +
                                "((t.CHIP_WT_GT)<>0) OR ((t.CHIP_VOL_CF)<>0) OR " +
                                "((t.BC_WT_GT)<>0) OR ((t.BC_VOL_CF)<>0))";

            }
            
            public static string UpdateTreeVolValWithMerchChipMarketValues(
                string p_strScenarioMerchChipMarketValuesTable,
                string p_strScenarioCostRevenueEscalatorValuesTable,
                string p_strScenarioId,
                string p_strTreeVolValTable)
            {
                return "UPDATE " + p_strTreeVolValTable + " t " +
                       "INNER JOIN (" + p_strScenarioMerchChipMarketValuesTable + " v " +
                                   "INNER JOIN " + p_strScenarioCostRevenueEscalatorValuesTable + " e " + 
                                   "ON v.scenario_id = e.scenario_id) " + 
                       "ON t.species_group = v.species_group AND t.diam_group=v.diam_group " + 
                       "SET t.merch_val_dpa = " + 
                                "IIF(t.RXCycle='1',t.merch_vol_cf * v.merch_value," + 
                                "IIF(t.RXCycle='2',t.merch_vol_cf * (v.merch_value * e.EscalatorMerchWoodRevenue_Cycle2)," + 
                                "IIF(t.RXCycle='3',t.merch_vol_cf * (v.merch_value * e.EscalatorMerchWoodRevenue_Cycle3)," + 
                                "IIF(t.RXCycle='4',t.merch_vol_cf * (v.merch_value * e.EscalatorMerchWoodRevenue_Cycle4),0))))," + 
                           "t.chip_mkt_val_pgt = " +
                                "IIF(t.RXCycle='1',v.chip_value," +
                                "IIF(t.RXCycle='2',(v.chip_value * e.EscalatorEnergyWoodRevenue_Cycle2)," +
                                "IIF(t.RXCycle='3',(v.chip_value * e.EscalatorEnergyWoodRevenue_Cycle3)," +
                                "IIF(t.RXCycle='4',(v.chip_value * e.EscalatorEnergyWoodRevenue_Cycle4),0))))," + 
                           "t.chip_val_dpa = " + 
                                "IIF(t.RXCycle='1',t.chip_wt_gt * v.chip_value," + 
                                "IIF(t.RXCycle='2',t.chip_wt_gt * (v.chip_value * e.EscalatorEnergyWoodRevenue_Cycle2)," + 
                                "IIF(t.RXCycle='3',t.chip_wt_gt * (v.chip_value * e.EscalatorEnergyWoodRevenue_Cycle3)," + 
                                "IIF(t.RXCycle='4',t.chip_wt_gt * (v.chip_value * e.EscalatorEnergyWoodRevenue_Cycle4),0))))," + 
                          "merch_to_chipbin_YN = " + 
                                "IIF(v.wood_bin='M','N','Y') " + 
                       "WHERE TRIM(UCASE(v.scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "'";

            }
           
           
            public static string PopulateFRCSInputTable(string p_strDestTableName,string p_strSourceTableName)
            {
                return "SELECT o.biosum_cond_id + o.rxpackage + o.rx + o.rxcycle AS Stand," +
                                "o.Slope AS [Percent Slope]," +
                                "o.gis_yard_dist AS [One-way Yarding Distance]," +
                                "Null AS BLANK," +
                                "o.elev AS [Project Elevation]," +
                                "o.[Harvesting system] AS [Harvesting System]," +
                                "o.[CHIPS TPA] AS [Chip tree per acre]," +
                                "o.[CHIPS Chip Fraction] AS [Residue fraction for chip trees]," +
                                "o.[CHIPS Average Vol (ft3)] AS [Chip trees average volume(ft3)]," +
                                "o.[CHIPS Average Density (lbs/ft3)]," +
                                "o.[CHIPS Hwd Proportion]," +
                                "o.[SMLOGS TPA] AS [Small log trees per acre]," +
                                "o.[SMLOGS Chip Fraction] AS [Small log trees residue fraction]," +
                                "o.[SMLOGS Average Vol (ft3)] AS [Small log trees average volume(ft3)]," +
                                "o.[SMLOGS Average Density (lbs/ft3)] AS [Small log trees average density(lbs/ft3)]," +
                                "o.[SMLOGS Hwd Proportion] AS [Small log trees hardwood proportion]," +
                                "o.[LGLOGS TPA] AS [Large log trees per acre]," +
                                "o.[LGLOGS Chip Fraction] AS [Large log trees residue fraction]," +
                                "o.[LGLOGS Average Vol (ft3)] AS [Large log trees average vol(ft3)]," +
                                "o.[LGLOGS Average Density (lbs/ft3)] AS [Large log trees average density(lbs/ft3)]," +
                                "o.[LGLOGS Hwd Proportion] AS [Large log trees hardwood proportion]," +
                                "Null AS BLANK1," +
                                "Null AS BLANK2," +
                                "Null AS BLANK3," +
                                "Null AS BLANK4," +
                                "Null AS BLANK5," +
                                "Null AS [Costs per green ton]," +
                                "Null AS [Costs per cubic foot]," +
                                "Null AS [Costs per cubic foot1]," +
                                "o.rxpackage + o.rx + o.rxcycle AS RxPackage_Rx_RxCycle " +
                        "INTO " + p_strDestTableName + " " +
                        "FROM " + p_strSourceTableName + " o ";
            }
            public static string AppendToFRCSInputTable(string p_strDestTableName, string p_strSourceTableName)
            {
                return "INSERT INTO " + p_strDestTableName + " " +
                                "(Stand," +
                                "[Percent Slope]," +
                                "[One-way Yarding Distance]," +
                                "BLANK," +
                                "[Project Elevation]," +
                                "[Harvesting System]," +
                                "[Chip tree per acre]," +
                                "[Residue fraction for chip trees]," +
                                "[Chip trees average volume(ft3)]," +
                                "[CHIPS Average Density (lbs/ft3)]," +
                                "[CHIPS Hwd Proportion]," +
                                "[Small log trees per acre]," +
                                "[Small log trees residue fraction]," +
                                "[Small log trees average volume(ft3)]," +
                                "[Small log trees average density(lbs/ft3)]," +
                                "[Small log trees hardwood proportion]," +
                                "[Large log trees per acre]," +
                                "[Large log trees residue fraction]," +
                                "[Large log trees average vol(ft3)]," +
                                "[Large log trees average density(lbs/ft3)]," +
                                "[Large log trees hardwood proportion]," +
                                "BLANK1," +
                                "BLANK2," +
                                "BLANK3," +
                                "BLANK4," +
                                "BLANK5," +
                                "[Costs per green ton]," +
                                "[Costs per cubic foot]," +
                                "[Costs per cubic foot1]," +
                                "RxPackage_Rx_RxCycle) " +
                        "SELECT Stand," +
                                "[Percent Slope]," +
                                "[One-way Yarding Distance]," +
                                "BLANK," +
                                "[Project Elevation]," +
                                "[Harvesting System]," +
                                "[Chip tree per acre]," +
                                "[Residue fraction for chip trees]," +
                                "[Chip trees average volume(ft3)]," +
                                "[CHIPS Average Density (lbs/ft3)]," +
                                "[CHIPS Hwd Proportion]," +
                                "[Small log trees per acre]," +
                                "[Small log trees residue fraction]," +
                                "[Small log trees average volume(ft3)]," +
                                "[Small log trees average density(lbs/ft3)]," +
                                "[Small log trees hardwood proportion]," +
                                "[Large log trees per acre]," +
                                "[Large log trees residue fraction]," +
                                "[Large log trees average vol(ft3)]," +
                                "[Large log trees average density(lbs/ft3)]," +
                                "[Large log trees hardwood proportion]," +
                                "BLANK1," +
                                "BLANK2," +
                                "BLANK3," +
                                "BLANK4," +
                                "BLANK5," +
                                "[Costs per green ton]," +
                                "[Costs per cubic foot]," +
                                "[Costs per cubic foot1]," +
                                "RxPackage_Rx_RxCycle " +
                        "FROM " + p_strSourceTableName; 
                               
            }
            public static string PopulateOPCOSTInputTable(string p_strDestTableName, string p_strSourceTableName)
            {
                return "SELECT o.biosum_cond_id + o.rxpackage + o.rx + o.rxcycle AS Stand," +
                                "o.Slope AS [Percent Slope]," +
                                "o.gis_yard_dist AS [One-way Yarding Distance]," +
                                "Null AS [YearCostCalc]," +
                                "o.elev AS [Project Elevation]," +
                                "o.[Harvesting system] AS [Harvesting System]," +
                                "o.[CHIPS TPA] AS [Chip tree per acre]," +
                                "o.[CHIPS Chip Fraction] AS [Residue fraction for chip trees]," +
                                "o.[CHIPS Average Vol (ft3)] AS [Chip trees average volume(ft3)]," +
                                "o.[CHIPS Average Density (lbs/ft3)]," +
                                "o.[CHIPS Hwd Proportion]," +
                                "o.[SMLOGS TPA] AS [Small log trees per acre]," +
                                "o.[SMLOGS Chip Fraction] AS [Small log trees residue fraction]," +
                                "o.[SMLOGS Average Vol (ft3)] AS [Small log trees average volume(ft3)]," +
                                "o.[SMLOGS Average Density (lbs/ft3)] AS [Small log trees average density(lbs/ft3)]," +
                                "o.[SMLOGS Hwd Proportion] AS [Small log trees hardwood proportion]," +
                                "o.[LGLOGS TPA] AS [Large log trees per acre]," +
                                "o.[LGLOGS Chip Fraction] AS [Large log trees residue fraction]," +
                                "o.[LGLOGS Average Vol (ft3)] AS [Large log trees average vol(ft3)]," +
                                "o.[LGLOGS Average Density (lbs/ft3)] AS [Large log trees average density(lbs/ft3)]," +
                                "o.[LGLOGS Hwd Proportion] AS [Large log trees hardwood proportion]," +
                                "o.[BRUSH CUT not utilized TPA] AS BrushCutTPA," +
                                "o.[BRUSH CUT not utilized Average Vol (ft3)] AS BrushCutAvgVol," +
                                "Null AS BLANK3," +
                                "Null AS BLANK4," +
                                "Null AS BLANK5," +
                                "Null AS [Costs per green ton]," +
                                "Null AS [Costs per cubic foot]," +
                                "Null AS [Costs per cubic foot1]," +
                                "o.rxpackage + o.rx + o.rxcycle AS RxPackage_Rx_RxCycle " +
                        "INTO " + p_strDestTableName + " " +
                        "FROM " + p_strSourceTableName + " o ";
            }
            public static string AppendToOPCOSTInputTable(string p_strDestTableName, string p_strSourceTableName)
            {
                return "INSERT INTO " + p_strDestTableName + " " +
                                "(Stand," +
                                "[Percent Slope]," +
                                "[One-way Yarding Distance]," +
                                "[YearCostCalc]," +
                                "[Project Elevation]," +
                                "[Harvesting System]," +
                                "[Chip tree per acre]," +
                                "[Residue fraction for chip trees]," +
                                "[Chip trees average volume(ft3)]," +
                                "[CHIPS Average Density (lbs/ft3)]," +
                                "[CHIPS Hwd Proportion]," +
                                "[Small log trees per acre]," +
                                "[Small log trees residue fraction]," +
                                "[Small log trees average volume(ft3)]," +
                                "[Small log trees average density(lbs/ft3)]," +
                                "[Small log trees hardwood proportion]," +
                                "[Large log trees per acre]," +
                                "[Large log trees residue fraction]," +
                                "[Large log trees average vol(ft3)]," +
                                "[Large log trees average density(lbs/ft3)]," +
                                "[Large log trees hardwood proportion]," +
                                "BrushCutTPA," +
                                "BrushCutAvgVol," +
                                "BLANK3," +
                                "BLANK4," +
                                "BLANK5," +
                                "[Costs per green ton]," +
                                "[Costs per cubic foot]," +
                                "[Costs per cubic foot1]," +
                                "RxPackage_Rx_RxCycle) " +
                        "SELECT Stand," +
                                "[Percent Slope]," +
                                "[One-way Yarding Distance]," +
                                "[YearCostCalc]," +
                                "[Project Elevation]," +
                                "[Harvesting System]," +
                                "[Chip tree per acre]," +
                                "[Residue fraction for chip trees]," +
                                "[Chip trees average volume(ft3)]," +
                                "[CHIPS Average Density (lbs/ft3)]," +
                                "[CHIPS Hwd Proportion]," +
                                "[Small log trees per acre]," +
                                "[Small log trees residue fraction]," +
                                "[Small log trees average volume(ft3)]," +
                                "[Small log trees average density(lbs/ft3)]," +
                                "[Small log trees hardwood proportion]," +
                                "[Large log trees per acre]," +
                                "[Large log trees residue fraction]," +
                                "[Large log trees average vol(ft3)]," +
                                "[Large log trees average density(lbs/ft3)]," +
                                "[Large log trees hardwood proportion]," +
                                "BrushCutTPA," +
                                "BrushCutAvgVol," +
                                "BLANK3," +
                                "BLANK4," +
                                "BLANK5," +
                                "[Costs per green ton]," +
                                "[Costs per cubic foot]," +
                                "[Costs per cubic foot1]," +
                                "RxPackage_Rx_RxCycle " +
                        "FROM " + p_strSourceTableName;

            }
            public static string PopulateFrcsVariableValuesTable(string p_strDestTableName,
                                                         string p_strSourceTableName,
                                                         string p_strSlopeExpr)
            {
                string sql = "";
                //removals per acre
                sql = "SELECT IIF(i.stand IS NOT NULL AND LEN(TRIM(i.stand)) " + 
                               ">= 25,MID(i.stand,1,25),i.stand) AS biosum_cond_id," + 
                            "MID(i.RxPackage_Rx_RxCycle,1,3) AS rxpackage," + 
                            "MID(i.RxPackage_Rx_RxCycle,4,3) AS rx," + 
                            "MID(i.RxPackage_Rx_RxCycle,7,1) AS rxcycle," + 
                            "i.[Harvesting System]," + 
                 			"IIF(i.[Chip tree per acre] IS NOT NULL,i.[chip tree per acre],0) " + 
                     "AS RemovalsCT,";
                     
                sql = sql + 
                                "IIF(i.[small log trees per acre] IS NOT NULL," + 
                                "i.[small log trees per acre],0) " + 
                                "AS RemovalsSLT,";
                     
                sql = sql + 
                                "IIF(i.[Large log trees per acre] IS NOT NULL," + 
                                "i.[Large log trees per acre],0) " + 
                                "AS RemovalsLLT,";
                     
                sql = sql + 
                                "RemovalsSLT + RemovalsLLT " + 
                                "AS RemovalsALT,";
                    
                sql = sql + 
                                "RemovalsCT + RemovalsSLT " + 
                                "AS RemovalsST,";
                     
                sql = sql + 
                                "RemovalsCT + RemovalsSLT + RemovalsLLT " + 
                                "AS Removals,";
                    
                //tree volume per cubic feet removed
                sql = sql + 
                                "IIF(i.[Chip trees average volume(ft3)] IS NOT NULL," + 
                                "i.[Chip trees average volume(ft3)],0) " + 
                                "AS TreeVolCT,";
                   
                sql = sql + 
    				            "IIF(i.[Large log trees average vol(ft3)] IS NOT NULL," + 
                                "i.[Large log trees average vol(ft3)],0) " + 
                                "AS TreeVolLLT,";
                   
                sql = sql + 
                                "IIF(i.[Small log trees average volume(ft3)] IS NOT NULL," + 
                                "i.[Small log trees average volume(ft3)],0) " + 
                                "AS TreeVolSLT,";
                   
                  
                    
                //volume per acre removed
                sql = sql + 
                                "RemovalsCT * TreeVolCT " + 
                                "AS VolPerAcreCT,";
                   
                sql = sql + 
                                "RemovalsSLT * TreeVolSLT " + 
                                "AS VolPerAcreSLT,";
                   
                sql = sql + 
    				            "RemovalsLLT * TreeVolLLT " + 
                                "AS VolPerAcreLLT,";
                   
                sql = sql + 
                                "VolPerAcreSLT + VolPerAcreLLT " + 
                                "AS VolPerAcreALT,";
                   
                sql = sql + 
                                "VolPerAcreCT + VolPerAcreSLT " + 
                               "AS VolPerAcreST,";
                   
                sql = sql + 
                               "(VolPerAcreCT + VolPerAcreSLT + VolPerAcreLLT) " + 
                               "AS VolPerAcre,";
                   
                   
                sql = sql + "IIF(RemovalsST > 0,VolPerAcreST/RemovalsST,0) AS TreeVolST,";
                sql = sql + "IIF(RemovalsALT > 0,VolPerAcreALT/RemovalsALT,0) AS TreeVolALT,";
                sql = sql + "IIF(Removals > 0,VolPerAcre/Removals,0) AS TreeVol,";
                sql = sql + "i.[Percent Slope] AS Slope ";
    	
                sql = sql + 
                                 "INTO " + p_strDestTableName + " " + 
                                 "FROM " + p_strSourceTableName + " i " + 
                                 "WHERE i.[Percent Slope] " + p_strSlopeExpr ;

                return sql;
            }
            /// <summary>
            /// Update the Warning Messages table with the FRCS output results.
            /// </summary>
            /// <param name="p_oHarvestMethodItem"></param>
            /// <param name="p_strUpdateFRCSWarningMessagesTableName"></param>
            /// <param name="p_strBinTableName"></param>
            /// <returns></returns>
            public static string UpdateFRCSWarningMessages(
                     uc_processor_scenario_run.FRCSHarvestMethodItem p_oHarvestMethodItem,
                     string p_strUpdateFRCSWarningMessagesTableName, string p_strBinTableName)
            {
                string strSQL = "";
                strSQL = "";
                //All Logs
                strSQL = "a.alllogvol=";
                if (p_oHarvestMethodItem.MaxCubicFootVolumeAllLogsDefault > 0)
                {
                    if (p_oHarvestMethodItem.MaxCubicFootVolumeAllLogs >
                        p_oHarvestMethodItem.MaxCubicFootVolumeAllLogsDefault)
                    {
                        strSQL = strSQL + "'ALLLOGVOL=" +
                                            p_oHarvestMethodItem.MaxCubicFootVolumeAllLogsDefault.ToString().Trim() +
                                            "/" +
                                            p_oHarvestMethodItem.MaxCubicFootVolumeAllLogs.ToString().Trim() +
                                          "'";
                    }
                    else
                    {
                        strSQL = strSQL + "''";
                    }
                }
                else
                {
                    strSQL = strSQL + "''";
                }
                //All Trees
                strSQL = strSQL + ",a.allvol=";
                if (p_oHarvestMethodItem.MaxCubicFootVolumeAllTreesDefault > 0)
                {
                    if (p_oHarvestMethodItem.MaxCubicFootVolumeAllTrees >
                        p_oHarvestMethodItem.MaxCubicFootVolumeAllTreesDefault)
                    {
                        strSQL = strSQL + "'ALLVOL=" +
                                            p_oHarvestMethodItem.MaxCubicFootVolumeAllTreesDefault.ToString().Trim() +
                                            "/" +
                                            p_oHarvestMethodItem.MaxCubicFootVolumeAllTrees.ToString().Trim() +
                                          "'";


                    }
                    else
                    {
                        strSQL = strSQL + "''";
                    }
                }
                else
                {
                    strSQL = strSQL + "''";
                }
                //Small Log Volume
                strSQL = strSQL + ",a.smlogvol=";
                if (p_oHarvestMethodItem.MaxCubicFootVolumeSmLogsDefault > 0)
                {
                    if (p_oHarvestMethodItem.MaxCubicFootVolumeSmLogs >
                        p_oHarvestMethodItem.MaxCubicFootVolumeSmLogsDefault)
                    {
                        strSQL = strSQL + "'SMLOGVOL=" +
                                            p_oHarvestMethodItem.MaxCubicFootVolumeSmLogsDefault.ToString().Trim() +
                                            "/" +
                                            p_oHarvestMethodItem.MaxCubicFootVolumeSmLogs.ToString().Trim() +
                                          "'";


                    }
                    else
                    {
                        strSQL = strSQL + "''";
                    }
                }
                else
                {
                    strSQL = strSQL + "''";
                }
                //large Log Volume
                strSQL = strSQL + ",a.lglogvol=";
                if (p_oHarvestMethodItem.MaxCubicFootVolumeLgLogsDefault > 0)
                {
                    if (p_oHarvestMethodItem.MaxCubicFootVolumeLgLogs >
                        p_oHarvestMethodItem.MaxCubicFootVolumeLgLogsDefault)
                    {
                        strSQL = strSQL + "'LGLOGVOL=" +
                                            p_oHarvestMethodItem.MaxCubicFootVolumeLgLogsDefault.ToString().Trim() +
                                            "/" +
                                            p_oHarvestMethodItem.MaxCubicFootVolumeLgLogs.ToString().Trim() +
                                          "'";


                    }
                    else
                    {
                        strSQL = strSQL + "''";
                    }
                }
                else
                {
                    strSQL = strSQL + "''";
                }
                //Chip Volume
                strSQL = strSQL + ",a.chipvol=";
                if (p_oHarvestMethodItem.MaxCubicFootVolumeChipsDefault > 0)
                {
                    if (p_oHarvestMethodItem.MaxCubicFootVolumeChips >
                        p_oHarvestMethodItem.MaxCubicFootVolumeChipsDefault)
                    {
                        strSQL = strSQL + "'CHIPSVOL=" +
                                            p_oHarvestMethodItem.MaxCubicFootVolumeChipsDefault.ToString().Trim() +
                                            "/" +
                                            p_oHarvestMethodItem.MaxCubicFootVolumeChips.ToString().Trim() +
                                          "'";


                    }
                    else
                    {
                        strSQL = strSQL + "''";
                    }
                }
                else
                {
                    strSQL = strSQL + "''";
                }
                //Large Log Per Acre To All Log Per Acre
                strSQL = strSQL + ",a.lglogpatoallpa=";
                if (p_oHarvestMethodItem.MaxLgLogsToAllLogsPercentDefault > 0)
                {
                    if (p_oHarvestMethodItem.MaxLgLogsToAllLogsPercent >
                        p_oHarvestMethodItem.MaxLgLogsToAllLogsPercentDefault)
                    {
                        strSQL = strSQL + "'LGLOGPATOALLPA=" +
                                            p_oHarvestMethodItem.MaxLgLogsToAllLogsPercentDefault.ToString().Trim() +
                                            "/" +
                                            p_oHarvestMethodItem.MaxLgLogsToAllLogsPercent.ToString().Trim() +
                                          "'";


                    }
                    else
                    {
                        strSQL = strSQL + "''";
                    }
                }
                else
                {
                    strSQL = strSQL + "''";
                }
                //Large Log Per Acre
                strSQL = strSQL + ",a.lglogpa=";
                if (p_oHarvestMethodItem.MaxCubicFootVolumeLgLogsPerAcreDefault > 0)
                {
                    if (p_oHarvestMethodItem.MaxCubicFootVolumeLgLogsPerAcre >
                        p_oHarvestMethodItem.MaxCubicFootVolumeLgLogsPerAcreDefault)
                    {
                        strSQL = strSQL + "'LGLLOGPA=" +
                                            p_oHarvestMethodItem.MaxCubicFootVolumeLgLogsPerAcreDefault.ToString().Trim() +
                                            "/" +
                                            p_oHarvestMethodItem.MaxCubicFootVolumeLgLogsPerAcre.ToString().Trim() +
                                          "'";
                        

                    }
                    else
                    {
                        strSQL = strSQL + "''";
                    }
                }
                else
                {
                    strSQL = strSQL + "''";
                }
                //Slope
                strSQL = strSQL + ",a.slope=";
                if (p_oHarvestMethodItem.MaxSlopePercentDefault > 0)
                {
                    if (p_oHarvestMethodItem.MaxSlopePercent >
                        p_oHarvestMethodItem.MaxSlopePercentDefault)
                    {
                        strSQL = strSQL + "'SLOPE=" +
                                            p_oHarvestMethodItem.MaxSlopePercentDefault.ToString().Trim() +
                                            "/" +
                                            p_oHarvestMethodItem.MaxSlopePercent.ToString().Trim() +
                                          "'";


                    }
                    else
                    {
                        strSQL = strSQL + "''";
                    }
                }
                else
                {
                    strSQL = strSQL + "''";
                }
                               
                strSQL = "UPDATE " + p_strUpdateFRCSWarningMessagesTableName + " a " +
                                    "INNER JOIN " + p_strBinTableName + " b " +
                                    "ON a.biosum_cond_id = b.biosum_cond_id AND " +
                                        "a.RXPackage = b.RXPackage AND " +
                                        "a.RX = b.RX AND " +
                                        "a.RXCycle = b.RXCycle " +
                                    "SET " + strSQL + " " +
                                    "WHERE TRIM(b.[Harvesting system]) = '" +
                                            p_oHarvestMethodItem.HarvestMethod.Trim() + "'";

                return strSQL;

            }
            public static string AppendToOPCOSTHarvestCostsTable(string p_strOPCOSTOutputTableName,
                                                          string p_strOPCOSTIdealOutputTableName,
                                                          string p_strOPCOSTInputTableName,
                                                          string p_strHarvestCostsTableName,
                                                          string p_strDateTimeCreated)
            {
                return "INSERT INTO " + p_strHarvestCostsTableName + " " +
                    "(biosum_cond_id, RXPackage, RX, RXCycle, " +
                    "harvest_cpa, chip_cpa, assumed_movein_cpa, " +
                    "ideal_harvest_cpa,ideal_chip_cpa, ideal_assumed_movein_cpa, " +
                    "override_YN, DateTimeCreated )" +
                    "SELECT o.biosum_cond_id, o.RxPackage,o.RX,o.RXCycle, " +
                    "IIF (RIGHT(CSTR(o.harvest_cpa), 6) = '1.#INF', 0,o.harvest_cpa ), " +
                    "o.chip_cpa, o.assumed_movein_cpa, " +
                    "IIF (RIGHT(CSTR(i.harvest_cpa), 6) = '1.#INF', 0,i.harvest_cpa ), " +
                    "i.ideal_chip_cpa, i.ideal_assumed_movein_cpa, " +
                    "IIF(n.[Unadjusted One-way Yarding distance] > 0 OR n.[Unadjusted Small log trees per acre] > 0 " +
                    "OR n.[Unadjusted Small log trees average volume (ft3)] > 0 OR n.[Unadjusted Large log trees per acre] > 0 " +
                    "OR n.[Unadjusted Large log trees average vol(ft3)] >0, 'Y','N') , " +
                    "'" + p_strDateTimeCreated + "' AS DateTimeCreated " +
                    "from (" + p_strOPCOSTOutputTableName + " o " +
                    "INNER JOIN " + p_strOPCOSTInputTableName + " n ON (o.biosum_cond_id = n.biosum_cond_id) AND " +
                    "(o.rxPackage = n.rxPackage) AND (o.RX = n.RX) AND (o.RXCycle = n.rxCycle)) " +
                    "INNER JOIN " + p_strOPCOSTIdealOutputTableName + " i ON (i.biosum_cond_id = n.biosum_cond_id) AND " +
                    "(i.rxPackage = n.rxPackage) AND (i.RX = n.RX) AND (i.RXCycle = n.rxCycle) ";
            }
            public static string AppendToHarvestCostsTable(string p_strFRCSOutputTableName,
                                                           string p_strHarvestCostsTableName,
                                                           string p_strDateTimeCreated)
            {
                return "INSERT INTO " + p_strHarvestCostsTableName + " " +
                              "(biosum_cond_id,RXPackage,RX,RXCycle,harvest_cpa,DateTimeCreated ) " +
                               "SELECT MID(STAND,1,25) AS biosum_cond_id," +
                                       "MID(rxpackage_rx_rxcycle,1,3) AS RXPackage," +
                                       "MID(rxpackage_rx_rxcycle,4,3) AS RX," +
                                       "MID(rxpackage_rx_rxcycle,7,1) AS RXCycle," +
                                       "IIF([$/Acre] IS NOT NULL," +
                                       "IIF(LEN(TRIM([$/Acre])) > 0 AND " +
                                           "ASC([$/Acre]) > 64 AND " +
                                           "ASC([$/Acre]) < 122 OR " +
                                           "ASC([$/Acre]) = 32 OR " +
                                           "ASC([$/Acre]) = 36,null,CDBL([$/Acre])),null) AS harvest_cpa, " +
                                       "'" + p_strDateTimeCreated + "' AS DateTimeCreated " +
                                "FROM " + p_strFRCSOutputTableName;
            }
            public static string UpdateHarvestCostsTableWithCompleteCostsPerAcre(
                string p_strScenarioCostRevenueEscalatorValuesTableName,
                string p_strTotalAdditionalCostsTableName,
                string p_strHarvestCostsTableName,
                string p_strScenarioId)
            {
                return "UPDATE " + p_strHarvestCostsTableName + " h " +
                       "INNER JOIN " +  p_strTotalAdditionalCostsTableName + " a " + 
                       "ON h.biosum_cond_id = a.biosum_cond_id AND h.rx=a.rx," + 
                       "scenario_cost_revenue_escalators e " + 
                       "SET h.complete_cpa = " + 
                           "IIF(h.RXCycle='1',(h.harvest_cpa+a.complete_additional_cpa)," + 
                           "IIF(h.RXCycle='2',(h.harvest_cpa+a.complete_additional_cpa) * e.EscalatorOperatingCosts_Cycle2," + 
                           "IIF(h.RXCycle='3',(h.harvest_cpa+a.complete_additional_cpa) * e.EscalatorOperatingCosts_Cycle3," + 
                           "IIF(h.RXCycle='4',(h.harvest_cpa+a.complete_additional_cpa) * e.EscalatorOperatingCosts_Cycle4,0)))) " +
                       "WHERE h.harvest_cpa IS NOT NULL AND h.harvest_cpa > 0  AND " + 
                             "TRIM(UCASE(e.scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "'";

            }
            public static string UpdateHarvestCostsTableWithIdealCompleteCostsPerAcre(
                string p_strScenarioCostRevenueEscalatorValuesTableName,
                string p_strTotalAdditionalCostsTableName,
                string p_strHarvestCostsTableName,
                string p_strScenarioId)
            {
                return "UPDATE " + p_strHarvestCostsTableName + " h " +
                       "INNER JOIN " + p_strTotalAdditionalCostsTableName + " a " +
                       "ON h.biosum_cond_id = a.biosum_cond_id AND h.rx=a.rx," +
                       "scenario_cost_revenue_escalators e " +
                       "SET h.ideal_complete_cpa = " +
                           "IIF(h.RXCycle='1',(h.ideal_harvest_cpa+a.complete_additional_cpa)," +
                           "IIF(h.RXCycle='2',(h.ideal_harvest_cpa+a.complete_additional_cpa) * e.EscalatorOperatingCosts_Cycle2," +
                           "IIF(h.RXCycle='3',(h.ideal_harvest_cpa+a.complete_additional_cpa) * e.EscalatorOperatingCosts_Cycle3," +
                           "IIF(h.RXCycle='4',(h.ideal_harvest_cpa+a.complete_additional_cpa) * e.EscalatorOperatingCosts_Cycle4,0)))) " +
                       "WHERE h.ideal_harvest_cpa IS NOT NULL AND h.ideal_harvest_cpa > 0  AND " +
                       "TRIM(UCASE(e.scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "'";
            }
                  

           
           
        }
        
		public class Reference
		{
			
			public string m_strRefHarvestMethodTable="";

			private Queries _oQueries=null;	
			private bool _bLoadDataSources=true;
			string m_strSql="";

			public Reference()
			{
			}

			public Queries ReferenceQueries
			{
				get {return _oQueries;}
				set {_oQueries=value;}
			}
			public bool LoadDatasource
			{
				get {return _bLoadDataSources;}
				set {_bLoadDataSources=value;}
			}


			public void LoadDatasources()
			{
				m_strRefHarvestMethodTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("HARVEST METHODS");
				
			
				if (m_strRefHarvestMethodTable.Trim().Length == 0 && this._oQueries._strScenarioType!="core") 
				{
					
					MessageBox.Show("!!Could Not Locate Harvest Methods Reference Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}					
			
			}
			/// <summary>
			/// 
			/// </summary>
			/// <param name="p_strTable">Harvest Method reference table</param>
			/// <param name="p_strFields">comma-delimited list of field names</param>
			/// <param name="p_strSteepSlopeYN">comma-delimited list of 'Y','N', or 'Y','N'</param>
			/// <returns></returns>
			public static string HarvestMethodTable_Select(string p_strTable,string p_strFields,string p_strWhereExpression)
			{
				return "SELECT " + p_strFields + " FROM " + p_strTable + " WHERE " + p_strWhereExpression;
			}

		}
        public class Utilities
        {
            public Utilities()
			{
			}
            /// <summary>
            /// Assign a sequential unique row number to each row.
            /// </summary>
            /// <param name="p_strIntoTableName"></param>
            /// <param name="p_strTableName"></param>
            /// <param name="p_strColumnName">Each data item for the column name must be unique</param>
            /// <param name="p_strRowNumberColumnName"></param>
            /// <returns></returns>
            public static string AssignRowNumberToEachRow(string p_strIntoTableName,string p_strTableName,string p_strColumnName,string p_strRowNumberColumnName)
            {
                string strSQL = "";

               
                if (p_strIntoTableName.Trim().Length > 0)
                {
                    strSQL = "SELECT (SELECT COUNT(a." + p_strColumnName + ") " +
                                     "FROM " + p_strTableName + " a " +
                                     "WHERE a." + p_strColumnName + " <= b." + p_strColumnName + ") AS " + p_strRowNumberColumnName + "," +
                                     "b." + p_strColumnName + " " +
                             "INTO " + p_strIntoTableName + " " + 
                             "FROM " + p_strTableName + " b " + 
                             "ORDER BY b." + p_strColumnName;
                }
                else
                {
                    strSQL = "SELECT (SELECT COUNT(a." + p_strColumnName + ") " +
                                     "FROM " + p_strTableName + " a " +
                                     "WHERE a." + p_strColumnName + " <= b." + p_strColumnName + ") AS " + p_strRowNumberColumnName + "," +
                                     "b." + p_strColumnName + " " +
                             "FROM " + p_strTableName + " b " + 
                             "ORDER BY b." + p_strColumnName;
                }

                return strSQL;
            }

        }
	}
}
