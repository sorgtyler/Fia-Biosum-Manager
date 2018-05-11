using System;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for Tables.
	/// </summary>
	public class Tables
	{
		public Project m_oProject = new Project();
		public CoreScenarioResults m_oCoreScenarioResults = new CoreScenarioResults();
		public CoreScenarioRuleDefinitions m_oCoreScenarioRuleDef = new CoreScenarioRuleDefinitions();
		public FIAPlot m_oFIAPlot = new FIAPlot();
		public FVS m_oFvs = new FVS();
		public TravelTime m_oTravelTime = new TravelTime();
		public Processor m_oProcessor = new Processor();
		public Scenario m_oScenario = new Scenario();
		public Audit m_oAudit = new Audit();
		public Reference m_oReference = new Reference();
		public ProcessorScenarioRun m_oProcessorScenarioRun = new ProcessorScenarioRun();
		public ProcessorScenarioRuleDefinitions m_oProcessorScenarioRuleDefinitions = new ProcessorScenarioRuleDefinitions();
		
		
		public Tables()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public class Project
		{
			private string strSQL="";
			public string DefaultProjectTableDbFile {get {return @"db\project.mdb";}}
			public string DefaultProjectTableName {get {return "project";}}
			public string DefaultProjectNotesTableDbFile {get {return @"db\project.mdb";}}
			public string DefaultProjectNotesTableName {get {return "notes";}}
			public string DefaultProjectLinksDepositoryTableDbFile {get {return @"db\project.mdb";}}
			public string DefaultProjectLinksDepositoryTableName {get {return "links_depository";}}
			public string DefaultProjectLinksCategoryTableDbFile {get {return @"db\project.mdb";}}
			public string DefaultProjectLinksCategoryTableName {get {return "links_category";}}
			public string DefaultProjectUserConfigTableDbFile {get {return @"db\project.mdb";}}
			public string DefaultProjectUserConfigTableName {get {return "user_config";}}
			public string DefaultProjectCoreScenarioDatasourceTableDbFile {get {return @"db\project.mdb";}}
			public string DefaultProjectCoreScenarioDatasourceTableName {get {return "core_scenario_datasource";}}
			public string DefaultProjectCoreScenarioTableDbFile {get {return @"db\project.mdb";}}
			public string DefaultProjectCoreScenarioTableName {get {return "core_scenario";}}
			public string DefaultProjectProcessorScenarioDatasourceTableDbFile {get {return @"db\project.mdb";}}
			public string DefaultProjectProcessorScenarioDatasourceTableName {get {return "processor_scenario_datasource";}}
			public string DefaultProjectProcessorScenarioTableDbFile {get {return @"db\project.mdb";}}
			public string DefaultProjectProcessorScenarioTableName {get {return "processor_scenario";}}
			public string DefaultProjectContactsTableDbFile {get {return @"db\project.mdb";}}
			public string DefaultProjectContactsTableName {get {return "contacts";}}
			static public string DefaultProjectDatasourceTableDbFile {get {return @"db\project.mdb";}}
			static public string DefaultProjectDatasourceTableName {get {return "datasource";}}
			public string DefaultProjectTravelTimesFormTableDbFile {get {return @"db\project.mdb";}}
			public string DefaultProjectTravelTimesFormTableName {get {return "form_travel_times";}}
			public Project()
			{
			}
			public void CreateContactsTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateContactsTableSQL(p_strTableName));
				CreateContactsTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateContactsTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","name");
			}
			public string CreateContactsTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					    "name CHAR(50)," + 
					    "process CHAR(50)," + 
					    "work_phone CHAR(20)," + 
						"organization CHAR(75)," + 
						"street_addr CHAR(30)," + 
						"city CHAR(30)," + 
						"state CHAR(2)," + 
						"zip CHAR(9)," + 
						"email CHAR(50))";

			}
			public void CreateDatasourceTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateDatasourceTableSQL(p_strTableName));
			}
			public string CreateDatasourceTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"table_type CHAR(60)," + 
					"path CHAR(254)," + 
					"file CHAR(50)," + 
					"table_name CHAR(50))";
			
			}
			public void CreateTravelTimesFormTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateTravelTimesFormTableSQL(p_strTableName));
			}
			public string CreateTravelTimesFormTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" + 
					   "map_proj CHAR(50)," + 
					   "disconnected_roads BYTE," + 
					   "travel_mode BYTE," + 
					   "psite_layer CHAR(50)," + 
					   "plot_layer CHAR(50)," + 
					   "weighted_speed_layer CHAR(50)," + 
					   "travel_times_mdb_file CHAR(254)," + 
					   "plot_mdb_file CHAR(254)," + 
					   "psite_mdb_file CHAR(254)," + 
					   "travel_times_layer_folder CHAR(254)," + 
					   "maximum_travel_time CHAR(4)," + 
					   "records_to_process1 BYTE DEFAULT 0," + 
					   "records_to_process2 BYTE DEFAULT 0)";
			}
			public void CreateLinksCategoryTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateLinksCategoryTableSQL(p_strTableName));
				CreateLinksCategoryTableIndexes(p_oAdo,p_oConn,p_strTableName);

			}
			public void CreateLinksCategoryTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx","category");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx2","subcategory");

			}
			public string CreateLinksCategoryTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"category BYTE," + 
					"subcategory INTEGER," + 
					"subcategory_desc CHAR(50))";
			}
			public void CreateLinksDepositoryTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateLinksDepositoryTableSQL(p_strTableName));
				CreateLinksDepositoryTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}

			public void CreateLinksDepositoryTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx","category");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx2","subcategory");
			}
			public string CreateLinksDepositoryTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"category BYTE," + 
					"subcategory INTEGER," + 
					"link CHAR(100)," + 
                    "description CHAR(100)," + 
					"list_yn CHAR(1) DEFAULT 'Y')";
			}
			public void CreateProjectTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateProjectTableSQL(p_strTableName));
				CreateProjectTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateProjectTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","proj_id");
			}
			public string CreateProjectTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"proj_id CHAR(20)," + 
					"created_by CHAR(30)," + 
					"created_date DATETIME," + 
					"company CHAR(100)," + 
					"description MEMO," + 
					"notes MEMO," + 
					"shared_file CHAR(254)," + 
					"project_root_directory CHAR(254)," + 
					"application_version CHAR(11))";       //version 4.0
			}

			public void CreateUserConfigTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateUserConfigTableSQL(p_strTableName));
			}
			public string CreateUserConfigTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"user_name CHAR(50)," + 
					"personal_directory CHAR(254))";
			}


			public void CreateProjectLinksCategoryTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateProjectLinksCategoryTableSQL(p_strTableName));
				CreateProjectLinksCategoryTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateProjectLinksCategoryTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx","category");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx2","subcategory");
			}
			public string CreateProjectLinksCategoryTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"category BYTE," + 
					"subcategory INTEGER," +
					"subcategory_desc CHAR(50))";
			}
			public void CreateProjectLinksDepositoryTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateProjectLinksDepositoryTableSQL(p_strTableName));
				CreateProjectLinksCategoryTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateProjectLinksDepositoryTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx","category");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx2","subcategory");
			}

			public string CreateProjectLinksDepositoryTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"category BYTE," + 
					"subcategory INTEGER," +
					"link CHAR(100)," + 
					"description CHAR(100)," + 
					"list_yn CHAR(1) DEFAULT 'Y')";
			}
			public void CreateProjectNotesTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateProjectNotesTableSQL(p_strTableName));
			}
			public string CreateProjectNotesTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"notes MEMO)";
			}





		}
		public class CoreScenarioResults
		{
            //cycle1
			static public string DefaultScenarioResultsCycle1BestRxSummaryWithIntensityTableDbFile {get {return @"db\scenario_results.mdb";}}
			static public string DefaultScenarioResultsCycle1BestRxSummaryWithIntensityTableName {get {return "cycle1_best_rx_summary_with_intensity";}}
			static public string DefaultScenarioResultsCycle1BestRxSummaryTableDbFile {get {return @"db\scenario_results.mdb";}}
			static public string DefaultScenarioResultsCycle1BestRxSummaryTableName {get {return "cycle1_best_rx_summary";}}
            static public string DefaultScenarioResultsCycle1OptimizationTableDbFile { get { return @"db\scenario_results.mdb"; } }
            static public string DefaultScenarioResultsCycle1OptimizationTableName { get { return "cycle1_optimization"; } }
            static public string DefaultScenarioResultsTieBreakerTableDbFile { get { return @"db\scenario_results.mdb"; } }
            static public string DefaultScenarioResultsTieBreakerTableName { get { return "tiebreaker"; } }

            //other
            static public string DefaultScenarioResultsValidCombosFVSPrePostTableDbFile { get { return @"db\scenario_results.mdb"; } }
            static public string DefaultScenarioResultsValidCombosFVSPrePostTableName { get { return "validcombos_fvsprepost"; } }
            static public string DefaultScenarioResultsValidCombosFVSPreTableDbFile { get { return @"db\scenario_results.mdb"; } }
            static public string DefaultScenarioResultsValidCombosFVSPreTableName { get { return "validcombos_fvspre"; } }
            static public string DefaultScenarioResultsValidCombosFVSPostTableDbFile { get { return @"db\scenario_results.mdb"; } }
            static public string DefaultScenarioResultsValidCombosFVSPostTableName { get { return "validcombos_fvspost"; } }
            static public string DefaultScenarioResultsValidCombosTableDbFile { get { return @"db\scenario_results.mdb"; } }
            static public string DefaultScenarioResultsValidCombosTableName { get { return "validcombos"; } }
            static public string DefaultScenarioResultsCycle1EffectiveTableDbFile { get { return @"db\scenario_results.mdb"; } }
            static public string DefaultScenarioResultsCycle1EffectiveTableName { get { return "cycle1_effective"; } }
            static public string DefaultScenarioResultsTreeVolValSumTableDbFile { get { return @"db\scenario_results.mdb"; } }
            static public string DefaultScenarioResultsTreeVolValSumTableName { get { return "tree_vol_val_sum_by_rx"; } }
            static public string DefaultScenarioResultsTreeVolValSumByRxPackageTableDbFile { get { return @"db\scenario_results.mdb"; } }
            static public string DefaultScenarioResultsTreeVolValSumByRxPackageTableName { get { return "tree_vol_val_sum_by_rxpackage"; } }
            static public string DefaultScenarioResultsPlotRxCostRevenueVolumesTableName { get { return "stand_costs_revenue_volume_by_rx"; } }
            static public string DefaultScenarioResultsPlotRxPackageCostRevenueVolumesSumTableName { get { return  "stand_costs_revenue_volume_sum_by_rxpackage"; } }
            static public string DefaultScenarioResultsPSiteRxCostRevenueVolumesTableName { get { return "psite_AcreExpansion_costs_revenue_volume_by_rx"; } }
            static public string DefaultScenarioResultsPSiteRxPackageCostRevenueVolumesTableName { get { return "psite_AcreExpansion_costs_revenue_volume_by_rxpackage"; } }
            static public string DefaultScenarioResultsOwnerRxCostRevenueVolumesTableName { get { return "own_AcreExpansion_costs_revenue_volume_by_rx"; } }
            static public string DefaultScenarioResultsOwnerRxPackageCostRevenueVolumesTableName { get { return "own_AcreExpansion_costs_revenue_volume_by_rxpackage"; } }
            


			
			private string strSQL = "";
			public CoreScenarioResults()
			{
			}
			//
			//EFFECTIVE TABLE
			//
			public void CreateEffectiveTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateEffectiveTableSQL(p_strTableName));
				CreateEffectiveTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateEffectiveTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id,rxpackage,rx,rxcycle");
			}
			static public string CreateEffectiveTableSQL(string p_strTableName)
			{
				return   "CREATE TABLE " + p_strTableName + " (" +
			             "biosum_cond_id CHAR(25)," + 
                         "rxpackage CHAR(3)," +
		                 "rx CHAR(3)," + 
                         "rxcycle CHAR(1)," + 
						 "nr_dpa DOUBLE," + 
			             "pre_variable1_name CHAR(100)," + 
			             "post_variable1_name CHAR(100)," + 
  			             "pre_variable1_value DOUBLE," + 
			             "post_variable1_value DOUBLE," + 
					     "variable1_change DOUBLE," +
			             "variable1_better_yn CHAR(1)," + 
			             "variable1_worse_yn CHAR(1)," + 
			             "variable1_effective_yn CHAR(1)," + 
						 "pre_variable2_name CHAR(100)," + 
			             "post_variable2_name CHAR(100)," + 
						 "pre_variable2_value DOUBLE," + 
						 "post_variable2_value DOUBLE," + 
					     "variable2_change DOUBLE," + 
						 "variable2_better_yn CHAR(1)," + 
						 "variable2_worse_yn CHAR(1)," + 
						 "variable2_effective_yn CHAR(1)," + 
						 "pre_variable3_name CHAR(100)," + 
						 "post_variable3_name CHAR(100)," + 
						 "pre_variable3_value DOUBLE," + 
						 "post_variable3_value DOUBLE," +
					     "variable3_change DOUBLE," + 
						 "variable3_better_yn CHAR(1)," + 
						 "variable3_worse_yn CHAR(1)," + 
						 "variable3_effective_yn CHAR(1)," + 
						 "pre_variable4_name CHAR(100)," + 
						 "post_variable4_name CHAR(100)," + 
						 "pre_variable4_value DOUBLE," + 
						 "post_variable4_value DOUBLE," +
					     "variable4_change DOUBLE," + 
						 "variable4_better_yn CHAR(1)," + 
						 "variable4_worse_yn CHAR(1)," + 
						 "variable4_effective_yn CHAR(1)," + 
						 "overall_effective_yn CHAR(1))";
			}
            //
            //NEW EFFECTIVE TABLE
            //
            public void NewCreateEffectiveTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, NewCreateEffectiveTableSQL(p_strTableName));
                NewCreateEffectiveTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void NewCreateEffectiveTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "biosum_cond_id,rxpackage");
            }
            static public string NewCreateEffectiveTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                         "biosum_cond_id CHAR(25)," +
                         "rxpackage CHAR(3)," +
                         "RxCycle1 CHAR(3)," +
                         "cycle1_nr_dpa DOUBLE," +
                         "cycle1_pre_variable1_name CHAR(100)," +
                         "cycle1_post_variable1_name CHAR(100)," +
                         "cycle1_pre_variable1_value DOUBLE," +
                         "cycle1_post_variable1_value DOUBLE," +
                         "cycle1_variable1_change DOUBLE," +
                         "cycle1_variable1_better_yn CHAR(1)," +
                         "cycle1_variable1_worse_yn CHAR(1)," +
                         "cycle1_variable1_effective_yn CHAR(1)," +
                         "cycle1_pre_variable2_name CHAR(100)," +
                         "cycle1_post_variable2_name CHAR(100)," +
                         "cycle1_pre_variable2_value DOUBLE," +
                         "cycle1_post_variable2_value DOUBLE," +
                         "cycle1_variable2_change DOUBLE," +
                         "cycle1_variable2_better_yn CHAR(1)," +
                         "cycle1_variable2_worse_yn CHAR(1)," +
                         "cycle1_variable2_effective_yn CHAR(1)," +
                         "cycle1_pre_variable3_name CHAR(100)," +
                         "cycle1_post_variable3_name CHAR(100)," +
                         "cycle1_pre_variable3_value DOUBLE," +
                         "cycle1_post_variable3_value DOUBLE," +
                         "cycle1_variable3_change DOUBLE," +
                         "cycle1_variable3_better_yn CHAR(1)," +
                         "cycle1_variable3_worse_yn CHAR(1)," +
                         "cycle1_variable3_effective_yn CHAR(1)," +
                         "cycle1_pre_variable4_name CHAR(100)," +
                         "cycle1_post_variable4_name CHAR(100)," +
                         "cycle1_pre_variable4_value DOUBLE," +
                         "cycle1_post_variable4_value DOUBLE," +
                         "cycle1_variable4_change DOUBLE," +
                         "cycle1_variable4_better_yn CHAR(1)," +
                         "cycle1_variable4_worse_yn CHAR(1)," +
                         "cycle1_variable4_effective_yn CHAR(1)," +
                         "cycle1_overall_effective_yn CHAR(1)," +
                         "RxCycle2 CHAR(3)," +
                         "cycle2_nr_dpa DOUBLE," +
                         "cycle2_pre_variable1_name CHAR(100)," +
                         "cycle2_post_variable1_name CHAR(100)," +
                         "cycle2_pre_variable1_value DOUBLE," +
                         "cycle2_post_variable1_value DOUBLE," +
                         "cycle2_variable1_change DOUBLE," +
                         "cycle2_variable1_better_yn CHAR(1)," +
                         "cycle2_variable1_worse_yn CHAR(1)," +
                         "cycle2_variable1_effective_yn CHAR(1)," +
                         "cycle2_pre_variable2_name CHAR(100)," +
                         "cycle2_post_variable2_name CHAR(100)," +
                         "cycle2_pre_variable2_value DOUBLE," +
                         "cycle2_post_variable2_value DOUBLE," +
                         "cycle2_variable2_change DOUBLE," +
                         "cycle2_variable2_better_yn CHAR(1)," +
                         "cycle2_variable2_worse_yn CHAR(1)," +
                         "cycle2_variable2_effective_yn CHAR(1)," +
                         "cycle2_pre_variable3_name CHAR(100)," +
                         "cycle2_post_variable3_name CHAR(100)," +
                         "cycle2_pre_variable3_value DOUBLE," +
                         "cycle2_post_variable3_value DOUBLE," +
                         "cycle2_variable3_change DOUBLE," +
                         "cycle2_variable3_better_yn CHAR(1)," +
                         "cycle2_variable3_worse_yn CHAR(1)," +
                         "cycle2_variable3_effective_yn CHAR(1)," +
                         "cycle2_pre_variable4_name CHAR(100)," +
                         "cycle2_post_variable4_name CHAR(100)," +
                         "cycle2_pre_variable4_value DOUBLE," +
                         "cycle2_post_variable4_value DOUBLE," +
                         "cycle2_variable4_change DOUBLE," +
                         "cycle2_variable4_better_yn CHAR(1)," +
                         "cycle2_variable4_worse_yn CHAR(1)," +
                         "cycle2_variable4_effective_yn CHAR(1)," +
                         "cycle2_overall_effective_yn CHAR(1)," +
                         "RxCycle3 CHAR(3)," +
                         "cycle3_nr_dpa DOUBLE," +
                         "cycle3_pre_variable1_name CHAR(100)," +
                         "cycle3_post_variable1_name CHAR(100)," +
                         "cycle3_pre_variable1_value DOUBLE," +
                         "cycle3_post_variable1_value DOUBLE," +
                         "cycle3_variable1_change DOUBLE," +
                         "cycle3_variable1_better_yn CHAR(1)," +
                         "cycle3_variable1_worse_yn CHAR(1)," +
                         "cycle3_variable1_effective_yn CHAR(1)," +
                         "cycle3_pre_variable2_name CHAR(100)," +
                         "cycle3_post_variable2_name CHAR(100)," +
                         "cycle3_pre_variable2_value DOUBLE," +
                         "cycle3_post_variable2_value DOUBLE," +
                         "cycle3_variable2_change DOUBLE," +
                         "cycle3_variable2_better_yn CHAR(1)," +
                         "cycle3_variable2_worse_yn CHAR(1)," +
                         "cycle3_variable2_effective_yn CHAR(1)," +
                         "cycle3_pre_variable3_name CHAR(100)," +
                         "cycle3_post_variable3_name CHAR(100)," +
                         "cycle3_pre_variable3_value DOUBLE," +
                         "cycle3_post_variable3_value DOUBLE," +
                         "cycle3_variable3_change DOUBLE," +
                         "cycle3_variable3_better_yn CHAR(1)," +
                         "cycle3_variable3_worse_yn CHAR(1)," +
                         "cycle3_variable3_effective_yn CHAR(1)," +
                         "cycle3_pre_variable4_name CHAR(100)," +
                         "cycle3_post_variable4_name CHAR(100)," +
                         "cycle3_pre_variable4_value DOUBLE," +
                         "cycle3_post_variable4_value DOUBLE," +
                         "cycle3_variable4_change DOUBLE," +
                         "cycle3_variable4_better_yn CHAR(1)," +
                         "cycle3_variable4_worse_yn CHAR(1)," +
                         "cycle3_variable4_effective_yn CHAR(1)," +
                         "cycle3_overall_effective_yn CHAR(1)," +
                         "RxCycle4 CHAR(3)," +
                         "cycle4_nr_dpa DOUBLE," +
                         "cycle4_pre_variable1_name CHAR(100)," +
                         "cycle4_post_variable1_name CHAR(100)," +
                         "cycle4_pre_variable1_value DOUBLE," +
                         "cycle4_post_variable1_value DOUBLE," +
                         "cycle4_variable1_change DOUBLE," +
                         "cycle4_variable1_better_yn CHAR(1)," +
                         "cycle4_variable1_worse_yn CHAR(1)," +
                         "cycle4_variable1_effective_yn CHAR(1)," +
                         "cycle4_pre_variable2_name CHAR(100)," +
                         "cycle4_post_variable2_name CHAR(100)," +
                         "cycle4_pre_variable2_value DOUBLE," +
                         "cycle4_post_variable2_value DOUBLE," +
                         "cycle4_variable2_change DOUBLE," +
                         "cycle4_variable2_better_yn CHAR(1)," +
                         "cycle4_variable2_worse_yn CHAR(1)," +
                         "cycle4_variable2_effective_yn CHAR(1)," +
                         "cycle4_pre_variable3_name CHAR(100)," +
                         "cycle4_post_variable3_name CHAR(100)," +
                         "cycle4_pre_variable3_value DOUBLE," +
                         "cycle4_post_variable3_value DOUBLE," +
                         "cycle4_variable3_change DOUBLE," +
                         "cycle4_variable3_better_yn CHAR(1)," +
                         "cycle4_variable3_worse_yn CHAR(1)," +
                         "cycle4_variable3_effective_yn CHAR(1)," +
                         "cycle4_pre_variable4_name CHAR(100)," +
                         "cycle4_post_variable4_name CHAR(100)," +
                         "cycle4_pre_variable4_value DOUBLE," +
                         "cycle4_post_variable4_value DOUBLE," +
                         "cycle4_variable4_change DOUBLE," +
                         "cycle4_variable4_better_yn CHAR(1)," +
                         "cycle4_variable4_worse_yn CHAR(1)," +
                         "cycle4_variable4_effective_yn CHAR(1)," +
                         "cycle4_overall_effective_yn CHAR(1)," + 
                         "RxPackage_overall_effective_yn CHAR(1))";
            }
			//
			//TIE BREAKER TABLE
			//
			public void CreateTieBreakerTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateTieBreakerTableSQL(p_strTableName));
				CreateTieBreakerTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateTieBreakerTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id,rxpackage,rx,rxcycle");
			}
			static public string CreateTieBreakerTableSQL(string p_strTableName)
			{
				return   "CREATE TABLE " + p_strTableName + " (" +
					"biosum_cond_id CHAR(25)," + 
                    "rxpackage CHAR(3)," + 
                    "rx CHAR(3)," + 
                    "rxcycle CHAR(1)," + 
                    "rx_intensity INTEGER," + 
					"pre_variable1_name CHAR(100)," + 
					"post_variable1_name CHAR(100)," + 
					"pre_variable1_value DOUBLE," + 
					"post_variable1_value DOUBLE," + 
					"variable1_change DOUBLE)";
					
			}
			//
			//VALID COMBO TABLE
			//
			public void CreateValidComboTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateValidComboTableSQL(p_strTableName));
				CreateValidComboTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateValidComboTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id,rxpackage,rx,rxcycle");
			}
			static public string CreateValidComboTableSQL(string p_strTableName)
			{
				return   "CREATE TABLE " + p_strTableName + " (" +
					"biosum_cond_id text(25)," + 
                    "rxpackage text(3)," +
					"rx text(3)," + 
                    "rxcycle text(1))";
			}
			public void CreateValidComboFVSPostTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateValidComboFVSPostTableSQL(p_strTableName));
				CreateValidComboFVSPostTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateValidComboFVSPostTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id,rxpackage,rx,rxcycle");
			}
			static public string CreateValidComboFVSPostTableSQL(string p_strTableName)
			{
				return   "CREATE TABLE " + p_strTableName + " (" +
							"biosum_cond_id text(25)," + 
                            "rxpackage text(3)," +
					        "rx text(3)," +
                            "rxcycle text(1)," + 
					        "variable1_yn text(1)," + 
					        "variable2_yn text(1)," + 
					        "variable3_yn text(1)," + 
					        "variable4_yn text(1))";

			}
			public void CreateValidComboFVSPreTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateValidComboFVSPreTableSQL(p_strTableName));
			    CreateValidComboFVSPreTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateValidComboFVSPreTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id,rxpackage,rx,rxcycle");
			}
			static public string CreateValidComboFVSPreTableSQL(string p_strTableName)
			{
				return   "CREATE TABLE " + p_strTableName + " (" +
					"biosum_cond_id text(25)," + 
                    "rxpackage text(3)," + 
                    "rx text(3)," + 
                    "rxcycle text(1)," + 
					"variable1_yn text(1)," + 
					"variable2_yn text(1)," + 
					"variable3_yn text(1)," + 
					"variable4_yn text(1))";

			}
			public void CreateValidComboFVSPrePostTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateValidComboFVSPrePostTableSQL(p_strTableName));
				CreateValidComboFVSPrePostTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateValidComboFVSPrePostTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id,rxpackage,rx,rxcycle");
			}
			static public string CreateValidComboFVSPrePostTableSQL(string p_strTableName)
			{
				return   "CREATE TABLE " + p_strTableName + " (" +
					"biosum_cond_id text(25)," + 
                    "rxpackage text(3)," + 
					"rx text(3)," + 
                    "rxcycle text(1))";
			}
			//
			//BEST TREATMENT TABLE
			//
			public void CreateBestRxSummaryCycle1Table(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateBestRxSummaryCycle1TableSQL(p_strTableName));
				this.CreateBestRxSummaryCycle1TableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateBestRxSummaryCycle1TableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id,rx");
			}
			static public string CreateBestRxSummaryCycle1TableSQL(string p_strTableName)
			{
				return   "CREATE TABLE " + p_strTableName + " (" +
					"biosum_cond_id text(25)," + 
                    "rx text(3)," + 
                    "acres double," + 
					"owngrpcd INTEGER," + 
					"optimization_value DOUBLE," + 
					"tiebreaker_value DOUBLE," + 
					"rx_intensity INTEGER)";
			}
			public void CreateBestRxSummaryCycle1WithIntensityTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateBestRxSummaryCycle1WithIntensityTableSQL(p_strTableName));
				CreateBestRxSummaryCycle1WithIntensityTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateBestRxSummaryCycle1WithIntensityTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id,rx");
			}
			static public string CreateBestRxSummaryCycle1WithIntensityTableSQL(string p_strTableName)
			{
				return   "CREATE TABLE " + p_strTableName + " (" +
					"biosum_cond_id text(25)," + 
                    "rx text(3)," + 
					"acres double," + 
					"owngrpcd INTEGER," + 
					"rx_intensity INTEGER)";
			}
			public void CreateBestRxSummaryCycle1TieBreakerTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateBestRxSummaryCycle1TieBreakerTableSQL(p_strTableName));
				CreateBestRxSummaryCycle1TieBreakerTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateBestRxSummaryCycle1TieBreakerTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id,rx");
			}
			static public string CreateBestRxSummaryCycle1TieBreakerTableSQL(string p_strTableName)
			{
				return   "CREATE TABLE " + p_strTableName + " (" +
					"biosum_cond_id text(25)," + 
                    "rx text(3)," + 
					"acres double," + 
					"owngrpcd INTEGER," + 
					"optimization_value DOUBLE," + 
					"tiebreaker_value DOUBLE," + 
					"rx_intensity INTEGER)";
			}
            
           
			//
			//OPTIMIZATION VARIABLE
			//
			public void CreateOptimizationTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateOptimizationTableSQL(p_strTableName));
				CreateOptimizationTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateOptimizationTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id,rxpackage,rx,rxcycle");
			}
			static public string CreateOptimizationTableSQL(string p_strTableName)
			{
				return   "CREATE TABLE " + p_strTableName + " (" +
					"biosum_cond_id CHAR(25)," + 
                    "rxpackage CHAR(3)," + 
                    "rx CHAR(3)," + 
					"rxcycle CHAR(1)," + 
					"pre_variable_name CHAR(100)," + 
					"post_variable_name CHAR(100)," + 
					"pre_variable_value DOUBLE," + 
					"post_variable_value DOUBLE," + 
					"change_value DOUBLE)";
					
			}
			//
			//OPTIMIZATION VARIABLE PLOT
			//
			public void CreateOptimizationPlotTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateOptimizationPlotTableSQL(p_strTableName));
				CreateOptimizationPlotTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateOptimizationPlotTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id");
			}
			static public string CreateOptimizationPlotTableSQL(string p_strTableName)
			{
				return   "CREATE TABLE " + p_strTableName + " (" +
					"biosum_cond_id CHAR(25)," + 
					"owngrpcd INTEGER," + 
					"acres DOUBLE," + 
					"optimization_value DOUBLE," + 
					"merch_haul_cost_psite INTEGER," + 
					"merch_haul_cpa DOUBLE," + 
					"merch_vol_cf_pa DOUBLE," +
					"merch_dollars_val_dpa DOUBLE," + 
					"chip_haul_cost_psite INTEGER," + 
					"chip_haul_cpa DOUBLE," + 
					"chip_yield_gt_pa DOUBLE," + 
					"chip_dollars_val_dpa DOUBLE," + 
					"net_rev_dpa DOUBLE," + 
					"harv_costs_cpa DOUBLE," + 
					"haul_costs_cpa DOUBLE)";
			}
            //
            //OPTIMIZATION BEST RX SUMMARY FOR CYCLE 1 BY STAND
            //
            public void CreateBestRxStandOptimizationSummaryCycle1Table(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateBestRxStandOptimizationSummaryCycle1TableSQL(p_strTableName));
                CreateBestRxStandOptimizationSummaryCycle1TableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateBestRxStandOptimizationSummaryCycle1TableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "id");
                p_oAdo.AddAutoNumber(p_oConn, p_strTableName, "id");
            }
            static public string CreateBestRxStandOptimizationSummaryCycle1TableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "id LONG," +
                    "biosum_cond_id CHAR(25)," +
                    "rx CHAR(3)," +
                    "chip_yield_cf DOUBLE," +
                    "merch_yield_cf DOUBLE," +
                    "chip_yield_gt DOUBLE," +
                    "merch_yield_gt DOUBLE," +
                    "chip_val_dpa DOUBLE," +
                    "merch_val_dpa DOUBLE," +
                    "harvest_onsite_cpa INTEGER," +
                    "haul_chip_cpa DOUBLE," +
                    "haul_merch_cpa DOUBLE," +
                    "merch_chip_nr_dpa DOUBLE," +
                    "merch_nr_dpa DOUBLE," +
                    "usebiomass_YN CHAR(1)," +
                    "max_nr_dpa DOUBLE)";
            }
			//
			//OPTIMIZATION VARIABLE PSITE
			//
			public void CreateOptimizationPSiteTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateOptimizationPSiteTableSQL(p_strTableName));
				CreateOptimizationPSiteTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateOptimizationPSiteTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","psite_id");
			}
			static public string CreateOptimizationPSiteTableSQL(string p_strTableName)
			{
				return   "CREATE TABLE " + p_strTableName + " (" +
					"psite_id INTEGER," + 
					"acres_sum DOUBLE," + 
					"optimization_value_sum DOUBLE," + 
					"merch_haul_cost_psite INTEGER," + 
					"merch_haul_cost_sum DOUBLE," + 
					"merch_vol_cf_sum DOUBLE," +
					"merch_dollars_val_sum DOUBLE," + 
					"chip_haul_cost_psite INTEGER," + 
					"chip_haul_cost_sum DOUBLE," + 
					"chip_yield_gt_sum DOUBLE," + 
					"chip_dollars_val_sum DOUBLE," + 
					"net_rev_dollars_sum DOUBLE," + 
					"harv_costs_sum DOUBLE," + 
					"haul_costs_sum DOUBLE)";
			}
			public void CreateOptimizationPSiteSumTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateOptimizationPSiteSumTableSQL(p_strTableName));
				CreateOptimizationPSiteSumTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateOptimizationPSiteSumTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","psite_id");
			}
			static public string CreateOptimizationPSiteSumTableSQL(string p_strTableName)
			{
				return   "CREATE TABLE " + p_strTableName + " (" +
					"psite_id INTEGER," + 
					"biocd BYTE," + 
					"acres_sum DOUBLE," + 
					"optimization_value_sum DOUBLE," + 
					"merch_haul_cost_sum DOUBLE," + 
					"merch_vol_cf_sum DOUBLE," +
					"merch_dollars_val_sum DOUBLE," + 
					"chip_haul_cost_sum DOUBLE," + 
					"chip_yield_gt_sum DOUBLE," + 
					"chip_dollars_val_sum DOUBLE," + 
					"net_rev_dollars_sum DOUBLE," + 
					"harv_costs_sum DOUBLE," + 
					"haul_costs_sum DOUBLE)";
			}
			//
			//OPTIMIZATION VARIABLE OWNERSHIP
			//
			public void CreateOptimizationOwnershipTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateOptimizationOwnershipTableSQL(p_strTableName));
				CreateOptimizationOwnershipTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateOptimizationOwnershipTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","owngrpcd");
			}
			static public string CreateOptimizationOwnershipTableSQL(string p_strTableName)
			{
				return   "CREATE TABLE " + p_strTableName + " (" +
					"owngrpcd INTEGER," + 
					"acres_sum DOUBLE," + 
					"optimization_value_sum DOUBLE," + 
					"merch_haul_cost_psite INTEGER," + 
					"merch_haul_cost_sum DOUBLE," + 
					"merch_vol_cf_sum DOUBLE," +
					"merch_dollars_val_sum DOUBLE," + 
					"chip_haul_cost_psite INTEGER," + 
					"chip_haul_cost_sum DOUBLE," + 
					"chip_yield_gt_sum DOUBLE," + 
					"chip_dollars_val_sum DOUBLE," + 
					"net_rev_dollars_sum DOUBLE," + 
					"harv_costs_sum DOUBLE," + 
					"haul_costs_sum DOUBLE)";
			}
			public void CreateOptimizationOwnershipSumTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateOptimizationOwnershipSumTableSQL(p_strTableName));
				CreateOptimizationOwnershipSumTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateOptimizationOwnershipSumTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","owngrpcd");
			}
			static public string CreateOptimizationOwnershipSumTableSQL(string p_strTableName)
			{
				return   "CREATE TABLE " + p_strTableName + " (" +
					"owngrpcd INTEGER," + 
					"acres_sum DOUBLE," + 
					"optimization_value_sum DOUBLE," + 
					"merch_haul_cost_sum DOUBLE," + 
					"chip_haul_cost_sum DOUBLE," + 
					"merch_vol_cf_sum DOUBLE," +
					"chip_yield_gt_sum DOUBLE," + 
					"net_rev_dollars_sum DOUBLE," + 
					"merch_dollars_val_sum DOUBLE," + 
					"chip_dollars_val_sum DOUBLE," + 
					"harv_costs_sum DOUBLE," + 
					"haul_costs_sum DOUBLE)";
			}
			//
			//INTENSITY WORK TABLE
			//
			public void CreateIntensityWorkTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateIntensityWorkTableSQL(p_strTableName));
			}
			public void CreateIntensityWorkTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id,rxpackage,rx,rxcycle");
			}
			static public string CreateIntensityWorkTableSQL(string p_strTableName)
			{
				return   "CREATE TABLE " + p_strTableName + " (" +
					"biosum_cond_id CHAR(25)," + 
                    "rxpackage CHAR(3)," + 
                    "rx CHAR(3)," + 
					"rxcycle CHAR(1)," + 
					"number_value DOUBLE," + 
					"number_value2 DOUBLE," + 
					"min_intensity INTEGER)";
			}
			//
			//HAUL COST TABLE
			//
			public void CreateHaulCostTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateHaulCostTableSQL(p_strTableName));
				CreateHaulCostTableIndexes(p_oAdo,p_oConn,p_strTableName);


			}
			public void CreateHaulCostTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","haul_cost_id");
				p_oAdo.AddAutoNumber(p_oConn,p_strTableName,"haul_cost_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","psite_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx2","railhead_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx3","biosum_plot_id");

			}

			static public string CreateHaulCostTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"haul_cost_id LONG," + 
					"biosum_plot_id CHAR(24)," + 
					"railhead_id INTEGER," + 
					"psite_id INTEGER," + 
					"transfer_cost DOUBLE DEFAULT 0," + 
					"road_cost DOUBLE DEFAULT 0," + 
					"rail_cost DOUBLE DEFAULT 0," + 
					"total_haul_cost DOUBLE DEFAULT 0," + 
					"materialcd CHAR(2))";
			}
			public void CreateHaulCostWorkTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateHaulCostWorkTableSQL(p_strTableName));
				CreateHaulCostWorkTableIndexes(p_oAdo,p_oConn,p_strTableName);


			}
			public void CreateHaulCostWorkTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","psite_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx2","railhead_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx3","biosum_plot_id");
			}

			static public string CreateHaulCostWorkTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"biosum_plot_id CHAR(24)," + 
					"railhead_id INTEGER," + 
					"psite_id INTEGER," + 
					"transfer_cost DOUBLE DEFAULT 0," + 
					"road_cost DOUBLE DEFAULT 0," + 
					"rail_cost DOUBLE DEFAULT 0," + 
					"total_haul_cost DOUBLE DEFAULT 0," + 
					"materialcd CHAR(2))";
			}
			public void CreateHaulCostRailroadTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateHaulCostTableSQL(p_strTableName));
				CreateHaulCostTableIndexes(p_oAdo,p_oConn,p_strTableName);


			}
			public void CreateHaulCostRailroadTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","psite_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx2","railhead_id");

			}

			static public string CreateHaulCostRailroadTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"railhead_id INTEGER," + 
					"psite_id INTEGER," + 
					"transfer_cost DOUBLE DEFAULT 0," + 
					"road_cost DOUBLE DEFAULT 0," + 
					"rail_cost DOUBLE DEFAULT 0," + 
					"total_haul_cost DOUBLE DEFAULT 0," + 
					"materialcd CHAR(2))";
			}
			//
			//TREE VOLUME AND VALUE SUM BY RX TABLE
			//
			public void CreateTreeVolValSumTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateTreeVolValSumTableSQL(p_strTableName));
				CreateTreeVolValSumTableIndexes(p_oAdo,p_oConn,p_strTableName);


			}
			public void CreateTreeVolValSumTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id,rxpackage,rx,rxcycle");
			}

			static public string CreateTreeVolValSumTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"biosum_cond_id CHAR(25)," + 
                    "rxpackage CHAR(3)," + 
                    "rx CHAR(3)," + 
					"rxcycle CHAR(1)," + 
					"chip_vol_cf DOUBLE," + 
					"chip_wt_gt DOUBLE," + 
					"chip_val_dpa DOUBLE," + 
					"merch_vol_cf DOUBLE," + 
					"merch_wt_gt DOUBLE," + 
					"merch_val_dpa DOUBLE)";
			}
            //
            //TREE VOLUME AND VALUE SUM BY RX PACKAGE TABLE
            //
            public void CreateTreeVolValSumByRxPackageTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateTreeVolValSumByRxPackageTableSQL(p_strTableName));
                CreateTreeVolValSumByRxPackageTableIndexes(p_oAdo, p_oConn, p_strTableName);


            }
            public void CreateTreeVolValSumByRxPackageTableIndexes(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "biosum_cond_id,rxpackage");
            }

            static public string CreateTreeVolValSumByRxPackageTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "chip_vol_cf DOUBLE," +
                    "chip_wt_gt DOUBLE," +
                    "chip_val_dpa DOUBLE," +
                    "merch_vol_cf DOUBLE," +
                    "merch_wt_gt DOUBLE," +
                    "merch_val_dpa DOUBLE)";
            }
			//
			//PRODUCT YIELDS NET REVENUE/COSTS SUMMARY TABLE
			//
			public void CreateProductYieldsTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateProductYieldsTableSQL(p_strTableName));
				CreateProductYieldsTableIndexes(p_oAdo,p_oConn,p_strTableName);


			}
			public void CreateProductYieldsTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id,rxpackage,rx,rxcycle");
			}

			static public string CreateProductYieldsTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"biosum_cond_id CHAR(25)," + 
                    "rxpackage CHAR(3)," + 
                    "rx CHAR(3)," + 
					"rxcycle CHAR(1)," + 
					"chip_yield_cf DOUBLE," + 
					"merch_yield_cf DOUBLE," + 
					"chip_yield_gt DOUBLE," + 
					"merch_yield_gt DOUBLE," + 
					"chip_val_dpa DOUBLE," + 
					"merch_val_dpa DOUBLE," + 
					"harvest_onsite_cpa DOUBLE," + 
					"haul_chip_cpa DOUBLE," + 
					"haul_merch_cpa DOUBLE," + 
					"merch_chip_nr_dpa DOUBLE," + 
					"merch_nr_dpa DOUBLE," + 
					"usebiomass_yn CHAR(1)," + 
					"max_nr_dpa DOUBLE)";
			}

            //
            //PRODUCT YIELDS NET REVENUE/COSTS SUMMARY BY PACKAGE TABLE
            //
            public void CreateProductYieldsByRxPackageTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateProductYieldsByRxPackageTableSQL(p_strTableName));
                CreateProductYieldsByRxPackageTableIndexes(p_oAdo, p_oConn, p_strTableName);


            }
            public void CreateProductYieldsByRxPackageTableIndexes(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "biosum_cond_id,rxpackage");
            }

            static public string CreateProductYieldsByRxPackageTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "chip_yield_cf DOUBLE," +
                    "merch_yield_cf DOUBLE," +
                    "chip_yield_gt DOUBLE," +
                    "merch_yield_gt DOUBLE," +
                    "chip_val_dpa DOUBLE," +
                    "merch_val_dpa DOUBLE," +
                    "harvest_onsite_cpa DOUBLE," +
                    "haul_chip_cpa DOUBLE," +
                    "haul_merch_cpa DOUBLE," +
                    "merch_chip_nr_dpa DOUBLE," +
                    "merch_nr_dpa DOUBLE," +
                    "max_nr_dpa DOUBLE)";
            }
            //
            //RX PLOT VALUES (COSTS, REVENUES, VOLUMES)
            //
            public void CreatePlotRxCostsRevenuesVolumesTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreatePlotRxCostsRevenuesVolumesTableSQL(p_strTableName));
                CreatePlotRxCostsRevenuesVolumesTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreatePlotRxCostsRevenuesVolumesTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "biosum_cond_id,rxpackage,rx,rxcycle");
            }
            static public string CreatePlotRxCostsRevenuesVolumesTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," + 
                    "rx CHAR(3)," + 
                    "rxcycle CHAR(1)," + 
                    "owngrpcd INTEGER," +
                    "acres DOUBLE," +
                    "merch_haul_cost_psite INTEGER," +
                    "merch_haul_cost_exp DOUBLE," +
                    "merch_vol_cf_exp DOUBLE," +
                    "merch_dollars_val_exp DOUBLE," +
                    "chip_haul_cost_psite INTEGER," +
                    "chip_haul_cost_exp DOUBLE," +
                    "chip_yield_gt_exp DOUBLE," +
                    "chip_dollars_val_exp DOUBLE," +
                    "net_rev_dollars_exp DOUBLE," +
                    "harv_costs_exp DOUBLE," +
                    "haul_costs_exp DOUBLE)";
            }
            //
            //RX PACKAGE PLOT VALUES (COSTS, REVENUES, VOLUMES)
            //
            public void CreatePlotRxPackageCostsRevenuesVolumesSumTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreatePlotRxPackageCostsRevenuesVolumesSumTableSQL(p_strTableName));
                CreatePlotRxPackageCostsRevenuesVolumesSumTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreatePlotRxPackageCostsRevenuesVolumesSumTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "biosum_cond_id,rxpackage");
            }
            static public string CreatePlotRxPackageCostsRevenuesVolumesSumTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "owngrpcd INTEGER," +
                    "acres DOUBLE," +
                    "merch_haul_cost_psite INTEGER," +
                    "merch_haul_cost_sum DOUBLE," +
                    "merch_vol_cf_sum DOUBLE," +
                    "merch_dollars_val_sum DOUBLE," +
                    "chip_haul_cost_psite INTEGER," +
                    "chip_haul_cost_sum DOUBLE," +
                    "chip_yield_gt_sum DOUBLE," +
                    "chip_dollars_val_sum DOUBLE," +
                    "net_rev_dollars_sum DOUBLE," +
                    "harv_costs_sum DOUBLE," +
                    "haul_costs_sum DOUBLE)";
            }
            //
            //TREATMENT PSITE TABLE
            //
            public void CreatePSiteRxCostsRevenuesVolumesTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreatePSiteRxCostsRevenuesVolumesTableSQL(p_strTableName));
                CreatePSiteRxCostsRevenuesVolumesTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreatePSiteRxCostsRevenuesVolumesTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "rxpackage,rx,rxcycle,psite_id");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "rxpackage,rx,rxcycle");
            }
            static public string CreatePSiteRxCostsRevenuesVolumesTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "psite_id INTEGER," +
                    "acres DOUBLE," +
                    "merch_haul_cost_psite INTEGER," +
                    "merch_haul_cost_sum DOUBLE," +
                    "merch_vol_cf_sum DOUBLE," +
                    "merch_dollars_val_sum DOUBLE," +
                    "chip_haul_cost_psite INTEGER," +
                    "chip_haul_cost_sum DOUBLE," +
                    "chip_yield_gt_sum DOUBLE," +
                    "chip_dollars_val_sum DOUBLE," +
                    "net_rev_dollars_sum DOUBLE," +
                    "harv_costs_sum DOUBLE," +
                    "haul_costs_sum DOUBLE)";
            }
            //
            //TREATMENT PSITE SUM TABLE
            //
            public void CreatePSiteRxCostsRevenuesVolumesSumTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreatePSiteRxCostsRevenuesVolumesSumTableSQL(p_strTableName));
                CreatePSiteRxCostsRevenuesVolumesSumTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreatePSiteRxCostsRevenuesVolumesSumTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "rxpackage,rx,rxcycle,psite_id");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "rxpackage,rx,rxcycle");
            }
            static public string CreatePSiteRxCostsRevenuesVolumesSumTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "psite_id INTEGER," +
                    "biocd BYTE," +
                    "acres_sum DOUBLE," +
                    "merch_haul_cost_sum DOUBLE," +
                    "merch_vol_cf_sum DOUBLE," +
                    "merch_dollars_val_sum DOUBLE," +
                    "chip_haul_cost_sum DOUBLE," +
                    "chip_yield_gt_sum DOUBLE," +
                    "chip_dollars_val_sum DOUBLE," +
                    "net_rev_dollars_sum DOUBLE," +
                    "harv_costs_sum DOUBLE," +
                    "haul_costs_sum DOUBLE)";
            }
            //
            //TREATMENT PACKAGE PSITE TABLE
            //
            public void CreatePSiteRxPackageCostsRevenuesVolumesTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreatePSiteRxPackageCostsRevenuesVolumesTableSQL(p_strTableName));
                CreatePSiteRxPackageCostsRevenuesVolumesTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreatePSiteRxPackageCostsRevenuesVolumesTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "rxpackage,psite_id");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "rxpackage");
            }
            static public string CreatePSiteRxPackageCostsRevenuesVolumesTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "rxpackage CHAR(3)," +
                    "psite_id INTEGER," +
                    "acres_sum DOUBLE," +
                    "merch_haul_cost_psite INTEGER," +
                    "merch_haul_cost_sum DOUBLE," +
                    "merch_vol_cf_sum DOUBLE," +
                    "merch_dollars_val_sum DOUBLE," +
                    "chip_haul_cost_psite INTEGER," +
                    "chip_haul_cost_sum DOUBLE," +
                    "chip_yield_gt_sum DOUBLE," +
                    "chip_dollars_val_sum DOUBLE," +
                    "net_rev_dollars_sum DOUBLE," +
                    "harv_costs_sum DOUBLE," +
                    "haul_costs_sum DOUBLE)";
            }

            //
            //TREATMENT OWNER TABLE
            //
            public void CreateOwnerRxCostsRevenuesVolumesTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateOwnerRxCostsRevenuesVolumesTableSQL(p_strTableName));
                CreateOwnerRxCostsRevenuesVolumesTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateOwnerRxCostsRevenuesVolumesTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "rxpackage,rx,rxcycle,owngrpcd");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "rxpackage,rx,rxcycle");
            }
            static public string CreateOwnerRxCostsRevenuesVolumesTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "owngrpcd INTEGER," +
                    "acres_sum DOUBLE," +
                    "merch_haul_cost_psite INTEGER," +
                    "merch_haul_cost_sum DOUBLE," +
                    "merch_vol_cf_sum DOUBLE," +
                    "merch_dollars_val_sum DOUBLE," +
                    "chip_haul_cost_psite INTEGER," +
                    "chip_haul_cost_sum DOUBLE," +
                    "chip_yield_gt_sum DOUBLE," +
                    "chip_dollars_val_sum DOUBLE," +
                    "net_rev_dollars_sum DOUBLE," +
                    "harv_costs_sum DOUBLE," +
                    "haul_costs_sum DOUBLE)";
            }
            //
            //TREATMENT OWNER SUM TABLE
            //
            public void CreateOwnerRxCostsRevenuesVolumesSumTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateOwnerRxCostsRevenuesVolumesSumTableSQL(p_strTableName));
                CreateOwnerRxCostsRevenuesVolumesSumTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateOwnerRxCostsRevenuesVolumesSumTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "rxpackage,rx,rxcycle,owngrpcd");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "rxpackage,rx,rxcycle");
            }
            static public string CreateOwnerRxCostsRevenuesVolumesSumTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "owngrpcd INTEGER," +
                    "acres DOUBLE," +
                    "merch_haul_cost_sum DOUBLE," +
                    "merch_vol_cf_sum DOUBLE," +
                    "merch_dollars_val_sum DOUBLE," +
                    "chip_haul_cost_sum DOUBLE," +
                    "chip_yield_gt_sum DOUBLE," +
                    "chip_dollars_val_sum DOUBLE," +
                    "net_rev_dollars_sum DOUBLE," +
                    "harv_costs_sum DOUBLE," +
                    "haul_costs_sum DOUBLE)";
            }
            //
            //TREATMENT PACKAGE OWNER TABLE
            //
            public void CreateOwnerRxPackageCostsRevenuesVolumesTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateOwnerRxPackageCostsRevenuesVolumesTableSQL(p_strTableName));
                CreateOwnerRxPackageCostsRevenuesVolumesTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateOwnerRxPackageCostsRevenuesVolumesTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "rxpackage,owngrpcd");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "rxpackage");
            }
            static public string CreateOwnerRxPackageCostsRevenuesVolumesTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "rxpackage CHAR(3)," +
                    "owngrpcd INTEGER," +
                    "acres DOUBLE," +
                    "merch_haul_cost_psite INTEGER," +
                    "merch_haul_cost_sum DOUBLE," +
                    "merch_vol_cf_sum DOUBLE," +
                    "merch_dollars_val_sum DOUBLE," +
                    "chip_haul_cost_psite INTEGER," +
                    "chip_haul_cost_sum DOUBLE," +
                    "chip_yield_gt_sum DOUBLE," +
                    "chip_dollars_val_sum DOUBLE," +
                    "net_rev_dollars_sum DOUBLE," +
                    "harv_costs_sum DOUBLE," +
                    "haul_costs_sum DOUBLE)";
            }





			
		
		}
		public class CoreScenarioRuleDefinitions
		{
			private string strSQL = "";
			static public string DefaultScenarioFvsVariablesTieBreakerTableDbFile {get {return @"core\db\scenario_core_rule_definitions.mdb";}}
			static public string DefaultScenarioFvsVariablesTieBreakerTableName {get {return "scenario_fvs_variables_tiebreaker";}}
			static public string DefaultScenarioFvsVariablesOptimizationTableDbFile {get {return @"core\db\scenario_core_rule_definitions.mdb";}}
            static public string DefaultScenarioFvsVariablesOptimizationTableName { get { return "scenario_fvs_variables_optimization"; } }
            static public string DefaultScenarioFvsVariablesOverallEffectiveTableDbFile { get { return @"core\db\scenario_core_rule_definitions.mdb"; } }
            static public string DefaultScenarioFvsVariablesOverallEffectiveTableName { get { return "scenario_fvs_variables_overall_effective"; } }
            static public string DefaultScenarioFvsVariablesTableDbFile { get { return @"core\db\scenario_core_rule_definitions.mdb"; } }
            static public string DefaultScenarioFvsVariablesTableName { get { return "scenario_fvs_variables"; } }
            static public string DefaultScenarioRxIntensityTableDbFile { get { return @"core\db\scenario_core_rule_definitions.mdb"; } }
            static public string DefaultScenarioRxIntensityTableName { get { return "scenario_rx_intensity"; } }
            static public string DefaultScenarioPSitesTableDbFile { get { return @"core\db\scenario_core_rule_definitions.mdb"; } }
            static public string DefaultScenarioPSitesTableName { get { return "scenario_psites"; } }
            static public string DefaultScenarioPlotFilterMiscTableDbFile { get { return @"core\db\scenario_core_rule_definitions.mdb"; } }
            static public string DefaultScenarioPlotFilterMiscTableName { get { return "scenario_plot_filter_misc"; } }
            static public string DefaultScenarioPlotFilterTableDbFile { get { return @"core\db\scenario_core_rule_definitions.mdb"; } }
            static public string DefaultScenarioPlotFilterTableName { get { return "scenario_plot_filter"; } }
            static public string DefaultScenarioMergeTableDbFile { get { return @"core\db\scenario_core_rule_definitions.mdb"; } }
            static public string DefaultScenarioMergeTableName { get { return "scenario_merge"; } }
            static public string DefaultScenarioLandOwnerGroupsTableDbFile { get { return @"core\db\scenario_core_rule_definitions.mdb"; } }
            static public string DefaultScenarioLandOwnerGroupsTableName { get { return "scenario_land_owner_groups"; } }
            static public string DefaultScenarioHarvestCostColumnsTableDbFile { get { return @"core\db\scenario_core_rule_definitions.mdb"; } }
            static public string DefaultScenarioHarvestCostColumnsTableName { get { return "scenario_harvest_cost_columns"; } }
            static public string DefaultScenarioDatasourceTableDbFile { get { return @"core\db\scenario_core_rule_definitions.mdb"; } }
            static public string DefaultScenarioDatasourceTableName { get { return @"scenario_datasource"; } }
            static public string DefaultScenarioCostsTableDbFile { get { return @"core\db\scenario_core_rule_definitions.mdb"; } }
            static public string DefaultScenarioCostsTableName { get { return "scenario_costs"; } }
            static public string DefaultScenarioTableDbFile { get { return @"core\db\scenario_core_rule_definitions.mdb"; } }
            static public string DefaultScenarioTableName { get { return "scenario"; } }
            static public string DefaultScenarioCondFilterMiscTableDbFile { get { return @"core\db\scenario_core_rule_definitions.mdb"; } }
            static public string DefaultScenarioCondFilterMiscTableName { get { return "scenario_cond_filter_misc"; } }
            static public string DefaultScenarioCondFilterTableDbFile { get { return @"core\db\scenario_core_rule_definitions.mdb"; } }
            static public string DefaultScenarioCondFilterTableName { get { return "scenario_cond_filter"; } }
            static public string DefaultScenarioProcessorScenarioSelectTableDbFile { get { return @"core\db\scenario_core_rule_definitions.mdb"; } }
            static public string DefaultScenarioProcessorScenarioSelectTableName { get { return "scenario_processor_scenario_select"; } }

			
			
			public CoreScenarioRuleDefinitions()
			{
			}
			
			public void CreateScenarioCostsTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateScenarioCostsTableSQL(p_strTableName));
				CreateScenarioCostsTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateScenarioCostsTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","scenario_id");
			}
            static public string CreateScenarioCostsTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," + 
					"chip_mkt_val_pgt DOUBLE DEFAULT 0," + 
					"road_haul_cost_pgt_per_hour DOUBLE DEFAULT 0," + 
					//"water_barring_roads_cpa DOUBLE," +
          //"brush_cutting_cpa DOUBLE," + 
					"rail_haul_cost_pgt_per_mile DOUBLE DEFAULT 0," + 
					"rail_chip_transfer_pgt_per_hour DOUBLE DEFAULT 0," + 
					"rail_merch_transfer_pgt_per_hour DOUBLE DEFAULT 0)";
			}
            public void CreateScenarioProcessorScenarioSelectTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateScenarioProcessorScenarioSelectTableSQL(p_strTableName));
                CreateScenarioProcessorScenarioSelectTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateScenarioProcessorScenarioSelectTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "scenario_id");
            }
            static public string CreateScenarioProcessorScenarioSelectTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "processor_scenario_id CHAR(20)," + 
                    "FullDetailsYN CHAR(1))";
            }
			
			
			public void CreateScenarioHarvestCostColumnsTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateScenarioHarvestCostColumnsTableSQL(p_strTableName));
                CreateScenarioHarvestCostColumnsTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateScenarioHarvestCostColumnsTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx","scenario_id");
			}
            static public string CreateScenarioHarvestCostColumnsTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," + 
					"ColumnName CHAR(50)," + 
					"Description CHAR(255))";
			}
			public void CreateScenarioLandOwnerGroupsTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateScenarioLandOwnerGroupsTableSQL(p_strTableName));
			    CreateScenarioLandOwnerGroupsTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateScenarioLandOwnerGroupsTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx","scenario_id");
			}
            static public string CreateScenarioLandOwnerGroupsTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," + 
					"owngrpcd INTEGER)";
			}
			public void CreateScenarioMergeTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateScenarioMergeTableSQL(p_strTableName));
				CreateScenarioMergeTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateScenarioMergeTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","cmborder");
			}
            static public string CreateScenarioMergeTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"cmborder BYTE," + 
					"mdbpathandfile CHAR(200))";
			}
			public void CreateScenarioPlotFilterTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateScenarioPlotFilterTableSQL(p_strTableName));
				CreateScenarioPlotFilterTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateScenarioPlotFilterTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx","scenario_id");
			}
            static public string CreateScenarioPlotFilterTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," + 
					"sql_command MEMO," + 
					"current_yn CHAR(1)," + 
					"table_list CHAR(200))";
			}

			public void CreateScenarioPlotFilterMiscTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateScenarioPlotFilterMiscTableSQL(p_strTableName));
				p_oAdo.AddIndex(p_oConn,"scenario_plot_filter_misc","scenario_plot_filter_misc_idx","scenario_id");
			}
			public void CreateScenarioPlotFilterMiscTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx","scenario_id");
			}
            static public string CreateScenarioPlotFilterMiscTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + "  (" +
					"scenario_id CHAR(20)," + 
					"yard_dist INTEGER," + 
					"yard_dist2 INTEGER)";
			}


			public void CreateScenarioCondFilterTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateScenarioCondFilterTableSQL(p_strTableName));
				CreateScenarioCondFilterTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateScenarioCondFilterTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx","scenario_id");
			}
            static public string CreateScenarioCondFilterTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," + 
					"sql_command MEMO," + 
					"current_yn CHAR(1)," + 
					"table_list CHAR(200))";
			}

			public void CreateScenarioCondFilterMiscTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateScenarioCondFilterMiscTableSQL(p_strTableName));
				p_oAdo.AddIndex(p_oConn,"scenario_cond_filter_misc","scenario_cond_filter_misc_idx","scenario_id");
			}
			public void CreateScenarioCondFilterMiscTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx","scenario_id");
			}
            static public string CreateScenarioCondFilterMiscTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + "  (" +
					"scenario_id CHAR(20)," + 
					"yard_dist INTEGER," + 
					"yard_dist2 INTEGER)";
			}


			public void CreateScenarioPSitesTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateScenarioPSitesTableSQL(p_strTableName));
				CreateScenarioPSitesTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateScenarioPSitesTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","scenario_id,psite_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx","scenario_id");
			}
            static public string CreateScenarioPSitesTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," + 
					"psite_id INTEGER," + 
					"name CHAR(30)," + 
					"trancd BYTE," + 
					"biocd BYTE," + 
					"selected_yn CHAR(1))";
			}
			public void CreateScenarioRxIntensityTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateScenarioRxIntensityTableSQL(p_strTableName));
				CreateScenarioRxIntensityTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateScenarioRxIntensityTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","scenario_id");
			}
            static public string CreateScenarioRxIntensityTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," + 
                   	"rx CHAR(3)," + 
					"rx_intensity INTEGER)";
			}

			public void CreateScenarioFVSVariablesTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
			    p_oAdo.SqlNonQuery(p_oConn,CreateScenarioFVSVariablesTableSQL(p_strTableName));
			}
            static public string CreateScenarioFVSVariablesTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," + 
                    "rxcycle CHAR(1)," + 
					"variable_number INTEGER," +
					"fvs_variables_list CHAR(255)," + 
					"pre_fvs_variable CHAR(100)," + 
					"post_fvs_variable CHAR(100)," + 
                    "better_expression MEMO," + 
                    "worse_expression MEMO," + 
                    "effective_expression MEMO," + 
					"current_yn CHAR(1))";
			}
			public void CreateScenarioFVSVariablesOverallEffectiveTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateScenarioFVSVariablesOverallEffectiveTableSQL(p_strTableName));
			}
            static public string CreateScenarioFVSVariablesOverallEffectiveTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," + 
                    "rxcycle CHAR(1)," + 
					"fvs_variables_list CHAR(255)," + 
					"overall_effective_expression MEMO," + 
					"nr_dpa_filter_enabled_yn CHAR(1)," + 
					"nr_dpa_filter_operator CHAR(2)," + 
					"nr_dpa_filter_value DOUBLE DEFAULT 0," + 
					"current_yn CHAR(1))";
			}
			//
			//core scenario rule definitions fvs variables optimization selection
			//
			public void CreateScenarioFVSVariablesOptimizationTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateScenarioFVSVariablesOptimizationTableSQL(p_strTableName));
			}
            static public string CreateScenarioFVSVariablesOptimizationTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," +
                    "rxcycle CHAR(1)," + 
					"optimization_variable CHAR(100)," +
					"fvs_variable_name CHAR(100)," + 
					"value_source CHAR(20)," + 
					"max_yn CHAR(1)," + 
					"min_yn CHAR(1)," + 
					"filter_enabled_yn CHAR(1)," + 
					"filter_operator CHAR(2)," + 
					"filter_value DOUBLE," + 
					"checked_yn CHAR(1)," + 
					"current_yn CHAR(1))";
			}
			public void CreateScenarioFVSVariablesTieBreakerTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn, CreateScenarioFVSVariablesTieBreakerTableSQL(p_strTableName));
			}
            static public string CreateScenarioFVSVariablesTieBreakerTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," +
                    "rxcycle CHAR(1)," + 
					"tiebreaker_method CHAR(100)," +
					"fvs_variable_name CHAR(100)," + 
					"value_source CHAR(20)," + 
					"max_yn CHAR(1)," + 
					"min_yn CHAR(1)," + 
					"checked_yn CHAR(1))";
			}





		}
		public class FVS
		{
            public static string[] g_strFVSOutTablesArray =  {"FVS_CASES",
															  "FVS_SUMMARY",
															  "FVS_SUMMARY_EAST",
                                                              "FVS_COMPUTE",
															  "FVS_TREELIST",
															  "FVS_ATRTLIST",
															  "FVS_CUTLIST",
															  "FVS_STRCLASS",
															  "FVS_POTFIRE",
															  "FVS_POTFIRE_EAST",
															  "FVS_CANPROFILE",
															  "FVS_FUELS",
															  "FVS_BURNREPORT",
															  "FVS_CONSUMPTION",
															  "FVS_MORTALITY",
															  "FVS_SNAGSUM",
															  "FVS_SNAGDET",
															  "FVS_CARBON",
															  "FVS_HRV_CARBON",
                                                              "FVS_DOWN_WOOD_COV",
                                                              "FVS_DOWN_WOOD_VOL",
															  "FVS_DM_SPP_SUM",
															  "FVS_DM_STND_SUM",
                                                              "FVS_DM_SZ_SUM",
															  "FVS_BM_MAIN",
															  "FVS_BM_BKP",
															  "FVS_BM_TREE",
                                                              "FVS_BM_VOL",
                                                              "FVS_ECONSUMMARY",
                                                              "FVS_ECONHARVESTVALUE"};

			static public string DefaultRxTableDbFile {get {return @"db\fvsmaster.mdb";}}
			static public string DefaultRxTableName {get {return "rx";}}

			static public string DefaultRxFvsCommandTableDbFile {get {return @"db\fvsmaster.mdb";}}
			static public string DefaultRxFvsCommandTableName {get {return "rx_fvs_commands";}}

			static public string DefaultRxHarvestCostColumnsTableDbFile {get {return @"db\fvsmaster.mdb";}}
			static public string DefaultRxHarvestCostColumnsTableName {get {return "rx_harvest_cost_columns";}}

			static public string DefaultRxPackageFvsCommandTableDbFile {get {return @"db\fvsmaster.mdb";}}
			static public string DefaultRxPackageFvsCommandTableName {get {return "rxpackage_fvs_commands";}}

			static public string DefaultRxPackageTableDbFile {get {return @"db\fvsmaster.mdb";}}
			static public string DefaultRxPackageTableName {get {return "rxpackage";}}

			static public string DefaultRxPackageMembersTableDbFile {get {return @"db\fvsmaster.mdb";}}
			static public string DefaultRxPackageMembersTableName {get {return "rxpackage_members";}}

			static public string DefaultRxPackageFvsCommandsOrderTableDbFile {get {return @"db\fvsmaster.mdb";}}
			static public string DefaultRxPackageFvsCommandsOrderTableName {get {return "rxpackage_fvs_commands_order";}}

			static public string DefaultFVSTreeTableDbFile {get {return @"processor\db\fvs_out_processor_in.mdb";}}
			static public string DefaultFVSTreeTableName {get {return "FVS_Tree";}}

            static public string DefaultOracleInputVolumesTable { get { return "biosum_volumes_input"; } }
            static public string DefaultOracleInputFCSVolumesTable { get { return "fcs_biosum_volumes_input"; } }

            static public string DefaultFVSTreeIdWorkTable { get { return "fvs_tree_id_work_table"; } }

            static public string DefaultFVSPrePostSeqNumTable { get { return "fvs_output_prepost_seqnum"; } }
            static public string DefaultFVSPrePostSeqNumTableDbFile { get { return @"db\fvsmaster.mdb"; } }

            static public string DefaultFVSPrePostSeqNumRxPackageAssgnTable { get { return "fvs_output_prepost_seqnum_rxpackage_assignment"; } }
            static public string DefaultFVSPrePostSeqNumRxPackageAssgnTableDbFile { get { return @"db\fvsmaster.mdb"; } }
			

			public FVS()
			{
			}

		    public static void CreateFVSCutListTable(ado_data_access p_oAdo)
		    {
		        p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, CreateFVSCutListTableSQL());
		    }

		    public static string CreateFVSCutListTableSQL()
		    {
		        return "CREATE TABLE FVS_CutList " +
		            "(CaseID CHAR(255)," +
		            "StandID CHAR(255)," +
                    "`Year` LONG," +
                    "PrdLen LONG," +
		            "TreeId CHAR(255)," +
                    "TreeIndex LONG," +
		            "Species CHAR(255)," +
                    "TreeVal LONG," +
                    "SSCD LONG," +
                    "PtIndex LONG," +
		            "TPA DOUBLE," +
		            "MortPA DOUBLE," +
		            "DBH DOUBLE," +
		            "DG DOUBLE," +
		            "Ht DOUBLE," +
		            "HtG DOUBLE," +
                    "PctCr LONG," +
		            "CrWidth DOUBLE," +
                    "MistCD LONG," +
		            "BAPctile DOUBLE," +
		            "PtBAL DOUBLE," +
		            "TCuFt DOUBLE," +
		            "MCuFt DOUBLE," +
		            "BdFt DOUBLE," +
                    "MDefect LONG," +
                    "BDefect LONG," +
                    "TruncHt LONG," +
		            "EstHt DOUBLE," +
                    "ActPt LONG," +
		            "Ht2TDCF SINGLE," +
		            "Ht2TDBF SINGLE," +
		            "TreeAge DOUBLE)";
		    }

			//
			//FVS_tree table
			//
			public void CreateFVSOutProcessorIn(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateFVSOutProcessorInTableSQL(p_strTableName));
				CreateFVSOutProcessorInTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateFVSOutProcessorInTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "id");
				p_oAdo.AddAutoNumber(p_oConn,p_strTableName,"id");		
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","fvs_tree_id");
			}
			public string CreateFVSOutProcessorInTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"id LONG," + 
					"biosum_cond_id CHAR(25)," + 
					"rxpackage CHAR(3)," + 
					"rx CHAR(3)," + 
					"rxcycle CHAR(1)," +
                    "rxyear CHAR(4)," + 
					"fvs_variant CHAR(2)," + 
					"cut_leave CHAR(1)," + 
					"fvs_species CHAR(6)," + 
					"tpa DOUBLE," + 
					"dbh DOUBLE," + 
					"ht DOUBLE," + 
                    "estht DOUBLE," + 
                    "pctcr DOUBLE," + 
                    "volcfnet DOUBLE," + 
					"volcfgrs DOUBLE," +
					"volcsgrs DOUBLE," + 
					"drybiom DOUBLE," + 
					"drybiot DOUBLE," + 
                    "voltsgrs DOUBLE," + 
					"fvs_tree_id CHAR(10)," + 
                    "FvsCreatedTree_YN CHAR(1) DEFAULT 'N'," +
                    "DateTimeCreated CHAR(22))";

			}
            //
            //FVS tree work table with cutlist.tcuft column table
            //
            public void CreateFVSOutTCuFt(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateFVSOutTCuFtTableSQL(p_strTableName));
                CreateFVSOutTCuFtTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateFVSOutTCuFtTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
               p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx", "fvs_tree_id");
            }
            public string CreateFVSOutTCuFtTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "rxyear CHAR(4)," +
                    "fvs_variant CHAR(2)," +
                    "TCuFt DOUBLE," + 
                    "fvs_tree_id CHAR(10))";
            }
            //
            //FVS_tree table audit
            //
            public void CreateFVSTreeIdAudit(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateFVSTreeIdAuditTableSQL(p_strTableName));
                CreateFVSTreeIdAuditTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateFVSTreeIdAuditTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "fvs_tree_id");
            }
            public string CreateFVSTreeIdAuditTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "id LONG," +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "rxyear CHAR(4)," +
                    "fvs_variant CHAR(2)," +
                    "fvs_tree_id CHAR(10)," + 
                    "Found_FvsTreeId_YN CHAR(1) DEFAULT 'N')";

            }
            //
            //FVS Cut list tree audit
            //
            //check that a single tree is only cut one time per package
            public void CreateFVSTreeIdCutAudit(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateFVSTreeIdCutAuditTableSQL(p_strTableName));
                CreateFVSTreeIdCutAuditTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateFVSTreeIdCutAuditTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "fvs_tree_id");
            }
            public string CreateFVSTreeIdCutAuditTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "fvs_tree_id CHAR(10)," +
                    "rxcycle1_YN CHAR(1) DEFAULT 'N'," +
                    "rxcycle2_YN CHAR(1) DEFAULT 'N'," +
                    "rxcycle3_YN CHAR(1) DEFAULT 'N'," +
                    "rxcycle4_YN CHAR(1) DEFAULT 'N'," + 
                    "Multiple_Cuts_YN CHAR(1) DEFAULT 'N')";
            }
            //
            //ORACLE biosum_volumes
            //
            public void CreateOracleInputBiosumVolumesTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateOracleInputBiosumVolumesTableSQL(p_strTableName));
                CreateOracleInputBiosumVolumesTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateOracleInputBiosumVolumesTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "id");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "fvs_tree_id");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "biosum_cond_id");
            }
            public string CreateOracleInputBiosumVolumesTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "id LONG," +
                    "biosum_cond_id CHAR(25)," +
                    "invyr INTEGER," +
                    "fvs_variant CHAR(2)," +
                    "spcd INTEGER," +
                    "dbh DOUBLE," +
                    "ht DOUBLE," +
                    "actualht DOUBLE," +
                    "vol_loc_grp CHAR(10)," +
                    "statuscd BYTE," +
                    "treeclcd BYTE," +
                    "cr DOUBLE," +
                    "cull DOUBLE," +
                    "roughcull DOUBLE," +
                    "decaycd BYTE," + 
                    "totage DOUBLE," +
                    "volcfnet DOUBLE," +
                    "volcfgrs DOUBLE," +
                    "volcsgrs DOUBLE," +
                    "drybiom DOUBLE," +
                    "drybiot DOUBLE," +
                    "voltsgrs DOUBLE," + 
                    "fvs_tree_id CHAR(10))";

            }
            public void CreateOracleInputFCSBiosumVolumesTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateOracleInputFCSBiosumVolumesTableSQL(p_strTableName));
                CreateOracleInputFCSBiosumVolumesTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateOracleInputFCSBiosumVolumesTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "TRE_CN");
               
            }
            public string CreateOracleInputFCSBiosumVolumesTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "STATECD INTEGER," + 
                    "COUNTYCD INTEGER," + 
                    "PLOT INTEGER," + 
                    "INVYR INTEGER," + 
                    "TREE INTEGER," + 
                    "VOL_LOC_GRP CHAR(10)," +
                    "SPCD INTEGER," + 
                    "DIA DOUBLE," + 
                    "HT DOUBLE," + 
                    "ACTUALHT DOUBLE," + 
                    "CR DOUBLE," + 
                    "STATUSCD BYTE," + 
                    "TREECLCD BYTE," + 
                    "ROUGHCULL DOUBLE," + 
                    "CULL DOUBLE," +
                    "DECAYCD BYTE," + 
                    "TOTAGE DOUBLE," +
                    "TRE_CN CHAR(34)," +
                    "CND_CN CHAR(34)," +
                    "PLT_CN CHAR(34)," +
                    "VOLCFGRS_CALC DOUBLE," + 
                    "VOLCSGRS_CALC DOUBLE," + 
                    "VOLCFNET_CALC DOUBLE," + 
                    "DRYBIOM_CALC DOUBLE," + 
                    "DRYBIOT_CALC DOUBLE," + 
                    "VOLTSGRS_CALC DOUBLE)";

            }
			//
			//RX table
			//
			/// <summary>
			/// Create the treatment table
			/// </summary>
			/// <param name="p_oAdo"></param>
			/// <param name="p_oConn"></param>
			/// <param name="p_strTableName"></param>
			public void CreateRxTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateRxTableSQL(p_strTableName));
				CreateRxTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateRxTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","rx");
			}
			public string CreateRxTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"rx CHAR(3)," + 
					"catid INTEGER," + 
					"subcatid INTEGER," + 
					"description MEMO," + 
					"HarvestMethodLowSlope CHAR(50)," + 
					"HarvestMethodSteepSlope CHAR(50))";
			}
			//
			//RX FVS Commands table
			//
			/// <summary>
			/// Create the table that defines the FVS commands associated with a treatment.
			/// </summary>
			/// <param name="p_oAdo"></param>
			/// <param name="p_oConn"></param>
			/// <param name="p_strTableName"></param>
			public void CreateRxFvsCommandsTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateRxFvsCommandsTableSQL(p_strTableName));
				CreateRxFvsCommandsTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateRxFvsCommandsTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","rx_fvscmd_index,rx,fvscmd");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","rx");
			}
			public string CreateRxFvsCommandsTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"rx_fvscmd_index LONG," + 
					"rx CHAR(3)," + 
					"fvscmd CHAR(30)," + 
					"fvscmd_id BYTE," + 
					"p1 CHAR(30)," + 
					"p2 CHAR(30)," + 
					"p3 CHAR(30)," + 
					"p4 CHAR(30)," + 
					"p5 CHAR(30)," + 
					"p6 CHAR(30)," + 
					"p7 CHAR(30)," + 
					"Other MEMO)"; 
					
			}
			//
			//RX HARVEST COST COLUMN
			//
			/// <summary>
			/// Create the rx harvest cost column table
			/// </summary>
			/// <param name="p_oAdo"></param>
			/// <param name="p_oConn"></param>
			/// <param name="p_strTableName"></param>
			public void CreateRxHarvestCostColumnTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateRxHarvestCostColumnTableSQL(p_strTableName));
				CreateRxHarvestCostColumnTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateRxHarvestCostColumnTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","rx,ColumnName");
                
			}
			static public string CreateRxHarvestCostColumnTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"rx CHAR(3)," + 
					"ColumnName CHAR(50)," + 
					"Description CHAR(255))";
			}
			//
			//RX PACKAGE TABLE
			//
			/// <summary>
			/// Create the Package table that defines an FVS treatment contained in a key/kcp file.
			/// </summary>
			/// <param name="p_oAdo"></param>
			/// <param name="p_oConn"></param>
			/// <param name="p_strTableName"></param>
			public void CreateRxPackageTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateRxPackageTableSQL(p_strTableName));
				CreateRxPackageTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateRxPackageTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "rxpackage");
			}
			public string CreateRxPackageTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"rxpackage CHAR(3)," + 
					"description MEMO," + 
					"rxcycle_length CHAR(2)," + 
					"simyear1_rx CHAR(3)," + 
					"simyear1_fvscycle CHAR(2)," + 
					"simyear2_rx CHAR(3)," + 
					"simyear2_fvscycle CHAR(2)," + 
					"simyear3_rx CHAR(3)," + 
					"simyear3_fvscycle CHAR(2)," + 
					"simyear4_rx CHAR(3)," + 
					"simyear4_fvscycle CHAR(2)," + 
					"kcpfile CHAR(254))";
			}
			//
			//RX PACKAGE FVS Commands table
			//
			/// <summary>
			/// Create the table that lists the FVS commands specific to a Package
			/// </summary>
			/// <param name="p_oAdo"></param>
			/// <param name="p_oConn"></param>
			/// <param name="p_strTableName"></param>
			public void CreateRxPackageFvsCommandsTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateRxPackageFvsCommandsTableSQL(p_strTableName));
				CreateRxPackageFvsCommandsTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateRxPackageFvsCommandsTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","rxpackage_fvscmd_index,rxpackage,fvscmd");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","rxpackage");
			}
			public string CreateRxPackageFvsCommandsTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"rxpackage_fvscmd_index LONG," + 
					"rxpackage CHAR(3)," + 
					"fvscmd CHAR(30)," + 
					"fvscmd_id BYTE," + 
					"list_index LONG," + 
					"p1 CHAR(30)," + 
					"p2 CHAR(30)," + 
					"p3 CHAR(30)," + 
					"p4 CHAR(30)," + 
					"p5 CHAR(30)," + 
					"p6 CHAR(30)," + 
					"p7 CHAR(30)," + 
					"Other MEMO)"; 
					
			}
			//
			//PACKAGE RX MEMBERS TABLE
			//
			/// <summary>
			/// Create the table that lists the RX members of a package. 
			/// </summary>
			/// <param name="p_oAdo"></param>
			/// <param name="p_oConn"></param>
			/// <param name="p_strTableName"></param>
			public void CreateRxPackageMembersTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateRxPackageMembersTableSQL(p_strTableName));
				CreateRxPackageMembersTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateRxPackageMembersTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","rxpackage");
			}
			public string CreateRxPackageMembersTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"rxpackage CHAR(3)," + 
					"rx CHAR(3))"; 
					
			}
			//
			//PACKAGE KCP FVS COMMANDS ORDER TABLE
			//
			/// <summary>
			/// Create the table that lists the order of the FVS Commands within a package
			/// </summary>
			/// <param name="p_oAdo"></param>
			/// <param name="p_oConn"></param>
			/// <param name="p_strTableName"></param>
			public void CreateRxPackageFvsCommandsOrderTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateRxPackageFvsCommandsOrderTableSQL(p_strTableName));
				CreateRxPackageFvsCommandsOrderTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateRxPackageFvsCommandsOrderTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","rxpackage");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx2","rx");
			}
			public string CreateRxPackageFvsCommandsOrderTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"rxpackage CHAR(3)," + 
					"rx CHAR(3)," + 
					"fvscycle CHAR(2)," + 
					"fvscmd CHAR(30)," + 
					"fvscmd_id BYTE)"; 
					
			}
			//
			//AUDIT TABLES
			//
			public static string CreateFVSOutputRxPrePostCycleYearTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (standid TEXT(255), pre_year1 INTEGER, post_year1 INTEGER," + 
					"pre_year2 INTEGER, post_year2 INTEGER ," + 
					"pre_year3 INTEGER, post_year3 INTEGER ," + 
					"pre_year4 INTEGER, post_year4 INTEGER)";
					
			}
			public static string CreateFVSOutputCycleYearTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (standid TEXT(255), `year` INTEGER,cycle INTEGER)";
			}
            //
            //FVS Tree Id used when creating FVS Text Input files
            //
            public static string CreateFVSTreeIdWorkTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "fvs_tree_id CHAR(9)," +
                    "fvs_variant CHAR(2)," +
                    "plotid CHAR(4)," +
                    "treeid CHAR(3)," +
                    "plottreeid CHAR(7))";

            }
            //
            //FVS Output PRE-POST Sequence Number Definitions
            //
            /// <summary>
            /// Create the table that defines the FVS Output PRE-POST Sequence Number Definitions.
            /// </summary>
            /// <param name="p_oAdo"></param>
            /// <param name="p_oConn"></param>
            /// <param name="p_strTableName"></param>
            public void CreateFVSOutputPrePostSeqNumTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateFVSOutputPrePostSeqNumTableSQL(p_strTableName));
                CreateFVSOutputPrePostSeqNumTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateFVSOutputPrePostSeqNumTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "PREPOST_SEQNUM_ID");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "tablename");
            }
            public string CreateFVSOutputPrePostSeqNumTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"PREPOST_SEQNUM_ID INTEGER," + 
					"TABLENAME CHAR(75)," + 
					"TYPE CHAR(1)," + 
					"RXCYCLE1_PRE_SEQNUM INTEGER," + 
					"RXCYCLE1_POST_SEQNUM INTEGER," + 
					"RXCYCLE2_PRE_SEQNUM INTEGER," + 
					"RXCYCLE2_POST_SEQNUM INTEGER," + 
					"RXCYCLE3_PRE_SEQNUM INTEGER," + 
					"RXCYCLE3_POST_SEQNUM INTEGER," + 
					"RXCYCLE4_PRE_SEQNUM INTEGER," + 
					"RXCYCLE4_POST_SEQNUM INTEGER," + 
					"RXCYCLE1_PRE_BASEYR_YN CHAR(1) DEFAULT 'N'," + 
					"RXCYCLE2_PRE_BASEYR_YN CHAR(1) DEFAULT 'N'," + 
					"RXCYCLE3_PRE_BASEYR_YN CHAR(1) DEFAULT 'N'," + 
					"RXCYCLE4_PRE_BASEYR_YN CHAR(1) DEFAULT 'N'," +
                    "RXCYCLE1_PRE_BEFORECUT_YN CHAR(1) DEFAULT 'N'," +
                    "RXCYCLE1_POST_BEFORECUT_YN CHAR(1) DEFAULT 'Y'," +
                    "RXCYCLE2_PRE_BEFORECUT_YN CHAR(1) DEFAULT 'N'," +
                    "RXCYCLE2_POST_BEFORECUT_YN CHAR(1) DEFAULT 'Y'," +
                    "RXCYCLE3_PRE_BEFORECUT_YN CHAR(1) DEFAULT 'N'," +
                    "RXCYCLE3_POST_BEFORECUT_YN CHAR(1) DEFAULT 'Y'," +
                    "RXCYCLE4_PRE_BEFORECUT_YN CHAR(1) DEFAULT 'N'," +
                    "RXCYCLE4_POST_BEFORECUT_YN CHAR(1) DEFAULT 'Y'," + 
                    "USE_SUMMARY_TABLE_SEQNUM_YN CHAR(1) DEFAULT 'N')"; 
					
			}
            //
            //FVS Output PRE-POST Sequence Number RX Package Assignments
            //
            /// <summary>
            /// Create the table that assigns seqnum to rx packages.
            /// </summary>
            /// <param name="p_oAdo"></param>
            /// <param name="p_oConn"></param>
            /// <param name="p_strTableName"></param>
            public void CreateFVSOutputPrePostSeqNumRxPackageAssgnTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateFVSOutputPrePostSeqNumRxPackageAssgnTableSQL(p_strTableName));
                CreateFVSOutputPrePostSeqNumRxPackageAssgnTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateFVSOutputPrePostSeqNumRxPackageAssgnTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "RXPACKAGE,PREPOST_SEQNUM_ID");
            }
            public string CreateFVSOutputPrePostSeqNumRxPackageAssgnTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                  "RXPACKAGE CHAR(3)," +
                    "PREPOST_SEQNUM_ID INTEGER)";

            }
            //
            //FVS Output PRE-POST Sequence Number Audit
            //
            public void CreateFVSOutputPrePostSeqNumAuditGenericTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateFVSOutputPrePostSeqNumAuditGenericTableSQL(p_strTableName));
                CreateFVSOutputPrePostSeqNumAuditGenericTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateFVSOutputPrePostSeqNumAuditGenericTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "SEQNUM,STANDID");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "STANDID,[YEAR]");
            }
            public string CreateFVSOutputPrePostSeqNumAuditGenericTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                  "SEQNUM INTEGER," +
                  "STANDID CHAR(255)," +
                  "[YEAR] INTEGER," +
                  "CYCLE1_PRE_YN CHAR(1)," +
                  "CYCLE1_POST_YN CHAR(1)," +
                  "CYCLE2_PRE_YN CHAR(1)," +
                  "CYCLE2_POST_YN CHAR(1)," +
                  "CYCLE3_PRE_YN CHAR(1)," +
                  "CYCLE3_POST_YN CHAR(1)," +
                  "CYCLE4_PRE_YN CHAR(1)," +
                  "CYCLE4_POST_YN CHAR(1))";

            }

            //
            //FVS Output StrClass PRE-POST Sequence Number Audit
            //
            public void CreateFVSOutputPrePostSeqNumAuditStrClassTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateFVSOutputPrePostSeqNumAuditStrClassTableSQL(p_strTableName));
                CreateFVSOutputPrePostSeqNumAuditStrClassTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateFVSOutputPrePostSeqNumAuditStrClassTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "SEQNUM,STANDID,REMOVAL_CODE");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "STANDID,[YEAR],REMOVAL_CODE");
            }
            public string CreateFVSOutputPrePostSeqNumAuditStrClassTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                  "SEQNUM INTEGER," +
                  "STANDID CHAR(255)," +
                  "[YEAR] INTEGER," +
                  "REMOVAL_CODE INTEGER," + 
                  "CYCLE1_PRE_YN CHAR(1)," +
                  "CYCLE1_POST_YN CHAR(1)," +
                  "CYCLE2_PRE_YN CHAR(1)," +
                  "CYCLE2_POST_YN CHAR(1)," +
                  "CYCLE3_PRE_YN CHAR(1)," +
                  "CYCLE3_POST_YN CHAR(1)," +
                  "CYCLE4_PRE_YN CHAR(1)," +
                  "CYCLE4_POST_YN CHAR(1))";

            }
           
            public static class Audit
            {
                public static class Post
                {

                    public static string CreateFVSPostAuditCutlistERROR_OUTPUTtableSQL(string p_strTableName)
                    {
                        return "CREATE TABLE " + p_strTableName + " (" +
                            "FVS_TREE_FILE CHAR(26)," +
                            "COLUMN_NAME CHAR(30)," +
                            "ERROR_DESC CHAR(60)," + 
                            "id LONG," +
                            "biosum_cond_id CHAR(25)," +
                            "rxpackage CHAR(3)," +
                            "rx CHAR(3)," +
                            "rxcycle CHAR(1)," +
                            "rxyear CHAR(4)," +
                            "fvs_variant CHAR(2)," +
                            "cut_leave CHAR(1)," +
                            "fvs_species CHAR(6)," +
                            "tpa DOUBLE," +
                            "dbh DOUBLE," +
                            "ht DOUBLE," +
                            "estht DOUBLE," +
                            "pctcr DOUBLE," +
                            "volcfnet DOUBLE," +
                            "volcfgrs DOUBLE," +
                            "volcsgrs DOUBLE," +
                            "drybiom DOUBLE," +
                            "drybiot DOUBLE," +
                            "voltsgrs DOUBLE," +
                            "fvs_tree_id CHAR(10)," +
                            "FvsCreatedTree_YN CHAR(1) DEFAULT 'N'," +
                            "DateTimeCreated CHAR(22))";

                    }
                    /// <summary>
                    /// Creates the POST-FVS audit table that DETAILS items in the BIOSUM FVS_TREE table
                    /// that are not found in other tables. For example, a BIOSUM_COND_ID in the 
                    /// BIOSUM FVS_TREE table should also exist in the BIOSUM CONDITION table. If it does not
                    /// then this table will be used to document the error.
                    /// </summary>
                    /// <param name="p_strTableName"></param>
                    /// <returns></returns>
                    public static string CreateFVSPostAuditCutlistNOTFOUND_ERRORtableSQL(string p_strTableName)
                    {
                        return "CREATE TABLE " + p_strTableName + " (" +
                            "FVS_TREE_FILE CHAR(26)," +
                            "COLUMN_NAME CHAR(30)," +
                            "NOTFOUND_VALUE CHAR(50)," +
                            "ERROR_DESC CHAR(60)," +
                            "id LONG," +
                            "biosum_cond_id CHAR(25)," +
                            "rxpackage CHAR(3)," +
                            "rx CHAR(3)," +
                            "rxcycle CHAR(1)," +
                            "rxyear CHAR(4)," +
                            "fvs_variant CHAR(2)," +
                            "cut_leave CHAR(1)," +
                            "fvs_species CHAR(6)," +
                            "tpa DOUBLE," +
                            "dbh DOUBLE," +
                            "ht DOUBLE," +
                            "estht DOUBLE," +
                            "pctcr DOUBLE," +
                            "volcfnet DOUBLE," +
                            "volcfgrs DOUBLE," +
                            "volcsgrs DOUBLE," +
                            "drybiom DOUBLE," +
                            "drybiot DOUBLE," +
                            "voltsgrs DOUBLE," +
                            "fvs_tree_id CHAR(10)," +
                            "FvsCreatedTree_YN CHAR(1) DEFAULT 'N'," +
                            "DateTimeCreated CHAR(22))";

                    }
                    /// <summary>
                    ///Creates the POST-FVS audit table that SUMMARIZES the validation of the BIOSUM FVS_TREE table data 
                    /// </summary>
                    /// <param name="p_strTableName"></param>
                    /// <returns></returns>
                    public static string CreateFVSPostAuditCutlistSUMMARYtableSQL(string p_strTableName)
                    {
                        return "CREATE TABLE " + p_strTableName + " (" +
                          "[INDEX] CHAR(3)," +
                          "FVS_TREE_FILE CHAR(26)," +
                          "COLUMN_NAME CHAR(30)," +
                          "NOVALUE_ERROR CHAR(5)," +
                          "NF_IN_COND_TABLE_ERROR CHAR(5)," +
                          "NF_IN_PLOT_TABLE_ERROR CHAR(5)," +
                          "VALUE_ERROR CHAR(20)," +
                          "NF_IN_RX_TABLE_ERROR CHAR(5)," +
                          "NF_RXPACKAGE_RXCYCLE_RX_ERROR CHAR(5)," +
                          "NF_IN_RXPACKAGE_TABLE_ERROR CHAR(5)," +
                          "NF_IN_TREE_TABLE_ERROR CHAR(5)," +
                          "TREE_SPECIES_CHANGE_WARNING CHAR(5)," +
                          "CREATED_DATE DATETIME)";
                    }
                    /// <summary>
                    ///Create the audit table used to check the tree data after appending FVS CUTLIST table data to the BIOSUM FVS_TREE table.
                    ///The purpose of the table is to contain matching FVS trees to FIA trees (by FVS_TREE_ID) to determine these items: 
                    ///1. Check if treatment cycle 1 FVS tree column data match FIA tree column data (ERROR item);
                    ///2. Check if the FVS tree species is different than the FIA tree species (WARNING item) 
                    /// </summary>
                    /// <param name="p_strTableName">Table name to create</param>
                    /// <param name="p_strDescriptionColumnName">Name of the column that will hold the warning or error description</param>
                    /// <returns></returns>
                    public static string CreateFVSPostAuditCutlistFVSFIA_TREEMATCHINGtableSQL(string p_strTableName,string p_strDescriptionColumnName)
                    {
                        return "CREATE TABLE " + p_strTableName + " (" +
                            "FVS_TREE_FILE CHAR(26)," +
                            "COLUMN_NAME CHAR(30)," +
                             p_strDescriptionColumnName + " CHAR(100)," + 
                            "ID LONG," +
                            "BIOSUM_COND_ID CHAR(25)," +
                            "RXCYCLE CHAR(1)," +
                            "FVS_TREE_FVS_TREE_ID CHAR(10)," +
                            "FIA_TREE_FVS_TREE_ID CHAR(10)," +
                            "FVS_TREE_SPCD INTEGER," + 
                            "FIA_TREE_SPCD INTEGER," + 
                            "FVS_TREE_DIA  SINGLE," + 
                            "FIA_TREE_DIA SINGLE," + 
                            "FVS_TREE_ESTHT DOUBLE," +
                            "FIA_TREE_ESTHT DOUBLE," +
                            "FVS_TREE_ACTUALHT DOUBLE," +
                            "FIA_TREE_ACTUALHT DOUBLE," +
                            "FVS_TREE_CR DOUBLE," +
                            "FIA_TREE_CR DOUBLE," +
                            "FVS_TREE_VOLCSGRS DOUBLE," +
                            "FIA_TREE_VOLCSGRS DOUBLE," +
                            "FVS_TREE_VOLCFGRS DOUBLE," +
                            "FIA_TREE_VOLCFGRS DOUBLE," + 
                            "FVS_TREE_VOLCFNET DOUBLE," +
                            "FIA_TREE_VOLCFNET DOUBLE," +
                            "FVS_TREE_VOLTSGRS DOUBLE," +
                            "FIA_TREE_VOLTSGRS DOUBLE," +
                            "FVS_TREE_DRYBIOT DOUBLE," +
                            "FIA_TREE_DRYBIOT DOUBLE," +
                            "FVS_TREE_DRYBIOM DOUBLE," +
                            "FIA_TREE_DRYBIOM DOUBLE," +
                            "FIA_TREE_STATUSCD BYTE," +
                            "FIA_TREE_TREECLCD BYTE," +
                            "FIA_TREE_CULL  DOUBLE," +
                            "FIA_TREE_ROUGHCULL DOUBLE," +
                            "FVSCREATEDTREE_YN CHAR(1) DEFAULT 'N')";

                   
                    }
                }
                public static class Pre
                {
                }
            }

            //To be used in fvs_input.cs for creating FVS_StandInit_WorkTable in the temporary biosum database, which will then be translated to FVS_StandInit
            public void CreateFVSInputStandInitTable(FIA_Biosum_Manager.ado_data_access p_oAdo,
                System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateFVSInputStandInitTableSQL(p_strTableName));
                CreateFVSInputStandInitTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }

            public void CreateFVSInputStandInitTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,
                System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "Stand_ID"); //TODO: Note that the Blank_Database used as an FVSIn.accdb template doesn't have primary keys.
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_Groups_idx", "Groups");
            }

            public string CreateFVSInputStandInitTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                       "Stand_ID CHAR(26)," + //indexed, duplicates okay //REQUIRED
                       "Variant CHAR(11)," + //REQUIRED
                       "Inv_Year LONG," + //REQUIRED
                       "Groups MEMO," + //indexed, duplicates okay
                       "AddFiles MEMO," +
                       "FVSKeywords MEMO," +
                       "Latitude DOUBLE," +
                       "Longitude DOUBLE," +
                       "Region LONG," +
                       "Forest LONG," +
                       "District LONG," +
                       "Compartment LONG," +
                       "Location LONG," + //RECOMMENDED
                       "Ecoregion CHAR(6)," +
                       "PV_Code CHAR(10)," +
                       "PV_Ref_Code LONG," +
                       "Age LONG," +
                       "Aspect DOUBLE," + //RECOMMENDED
                       "Slope DOUBLE," + //RECOMMENDED
                       "Elevation DOUBLE," +
                       "ElevFt DOUBLE," + //RECOMMENDED
                       "Basal_Area_Factor DOUBLE," + //REQUIRED
                       "Inv_Plot_Size DOUBLE," + //REQUIRED
                       "Brk_DBH DOUBLE," + //REQUIRED
                       "Num_Plots LONG," + //RECOMMENDED
                       "NonStk_Plots LONG," +
                       "Sam_Wt DOUBLE," +
                       "Stk_Pcnt DOUBLE," +
                       "DG_Trans LONG," +
                       "DG_Measure LONG," +
                       "HTG_Trans LONG," +
                       "HTG_Measure LONG," +
                       "Mort_Measure LONG," +
                       "Max_BA DOUBLE," +
                       "Max_SDI DOUBLE," +
                       "Site_Species CHAR(8)," +
                       "Site_Index DOUBLE," +
                       "Model_Type LONG," +
                       "Physio_Region LONG," +
                       "Forest_Type LONG," +
                       "State LONG," +
                       "County LONG," +
                       "Fuel_Model LONG," +
                       "Fuel_0_25_H DOUBLE," +
                       "Fuel_25_1_H DOUBLE," +
                       "Fuel_1_3_H DOUBLE," +
                       "Fuel_3_6_H DOUBLE," +
                       "Fuel_6_12_H DOUBLE," +
                       "Fuel_12_20_H DOUBLE," +
                       "Fuel_20_35_H DOUBLE," +
                       "Fuel_35_50_H DOUBLE," +
                       "Fuel_gt_50_H DOUBLE," +
                       "Fuel_0_25_S DOUBLE," +
                       "Fuel_25_1_S DOUBLE," +
                       "Fuel_1_3_S DOUBLE," +
                       "Fuel_3_6_S DOUBLE," +
                       "Fuel_6_12_S DOUBLE," +
                       "Fuel_12_20_S DOUBLE," +
                       "Fuel_20_35_S DOUBLE," +
                       "Fuel_35_50_S DOUBLE," +
                       "Fuel_gt_50_S DOUBLE," +
                       "Fuel_Litter DOUBLE," +
                       "Fuel_Duff DOUBLE," +
                       "Photo_Ref LONG," +
                       "Photo_code CHAR(13)" +
                       ")";
            }


            //To be used in fvs_input.cs for creating FVS_TreeInit_WorkTable in the temporary biosum database, which will then be translated to FVS_TreeInit
            public void CreateFVSInputTreeInitWorkTable(FIA_Biosum_Manager.ado_data_access p_oAdo,
                System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateFVSInputTreeInitWorkTableSQL(p_strTableName));
                CreateFVSInputTreeInitWorkTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }

            public void CreateFVSInputTreeInitWorkTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,
                System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "TreeInitID");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_Stand_ID_idx", "Stand_ID");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_StandPlot_ID_idx", "StandPlot_ID");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_Tree_ID_idx", "Tree_ID");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_PV_Code_idx", "PV_Code");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_TopoCode_idx", "TopoCode");
            }

            public string CreateFVSInputTreeInitWorkTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                       "TreeInitID AUTOINCREMENT, " + //Pkey and indexed
                       "Stand_ID CHAR(26)," + //indexed, duplicates okay //REQUIRED
                       "StandPlot_ID CHAR(40)," +
                    //"Plot_ID CHAR(4)," + //RECOMMENDED (how FVS defines this field)
                    //"Tree_ID CHAR(40)," + // (how FVS defines this field)
                       "Plot_ID DOUBLE," + //RECOMMENDED //This deviates from FVS DB documentation. Our current approach is to not use plot_IDs
                       "Tree_ID DOUBLE," + //This deviates from FVS DB documentation because we calculate Tree_ID as Tree.Subp*1000+Tree.Tree. Also indexed.
                       "Tree_Count DOUBLE," + //RECOMMENDED
                       "History DOUBLE," +
                       "Species CHAR(8)," + //REQUIRED
                       "DBH DOUBLE," + //REQUIRED
                       "DG DOUBLE," +
                       "Htcd DOUBLE," + //Temporary table only: If Htcd in {1,2,3}, then we import Ht and HtTopK from Biosum Ht and ActualHt
                       "Ht DOUBLE," + //RECOMMENDED
                       "HTG DOUBLE," +
                       "HtTopK DOUBLE," +
                       "CrRatio DOUBLE," + //RECOMMENDED
                       "Damage1 DOUBLE," +
                       "Severity1 DOUBLE," +
                       "Damage2 DOUBLE," +
                       "Severity2 DOUBLE," +
                       "Damage3 DOUBLE," +
                       "Severity3 DOUBLE," +
                       "TreeValue DOUBLE," +
                       "Prescription DOUBLE," +
                       "Age DOUBLE," +
                       "Slope DOUBLE," +
                       "Aspect DOUBLE," +
                       "PV_Code CHAR(10)," + //indexed, duplicates okay
                       "TopoCode DOUBLE," + //indexed, duplicates okay
                       "SitePrep DOUBLE" +

                       //To be used in the work table in the temp biosum mdb before the above records are inserted into FVS_TreeInit
                       ", fvs_dmg_ag1 CHAR(2)," +
                       "fvs_dmg_sv1 CHAR(2)," +
                       "fvs_dmg_ag2 CHAR(2)," +
                       "fvs_dmg_sv2 CHAR(2)," +
                       "fvs_dmg_ag3 CHAR(2)," +
                       "fvs_dmg_sv3 CHAR(2)," +
                       "mist_cl_cd LONG, " +
                       "hasBrokenTop BIT, " +
                       "cullbf DOUBLE,  " +
                       "TreeCN CHAR(34)  " +
                       ")";
            }


			
			


		}
		public class TravelTime
		{
			public string DefaultTravelTimeTableDbFile {get {return @"gis\db\gis_travel_times.mdb";}}
			public string DefaultTravelTimeTableName {get {return "travel_time";}}
			public string DefaultProcessingSiteTableDbFile {get {return @"gis\db\gis_travel_times.mdb";}}
			public string DefaultProcessingSiteTableName {get {return "processing_site";}}
			public string DefaultDisconnectedRoadTravelTimeOfZeroDbFile {get {return @"gis\db\gis_travel_times.mdb";}}
			public string DefaultDisconnectedRoadTravelTimeOfZeroTableName {get {return "disconnected_road_travel_time_of_zero";}}
			public string DefaultTravelTimeOfZeroDbFile {get {return @"gis\db\gis_travel_times.mdb";}}
			public string DefaultTravelTimeOfZeroTableName {get {return "travel_time_of_zero";}}
				

			public TravelTime()
			{
			}
			public void CreateDisconnectedRoadTravelTimeOfZeroTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateDisconnectedRoadTravelTimeOfZeroTableSQL(p_strTableName));
				CreateDisconnectedRoadTravelTimeOfZeroTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateDisconnectedRoadTravelTimeOfZeroTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","traveltime_id");
				p_oAdo.AddAutoNumber(p_oConn,p_strTableName,"traveltime_id");				
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","psite_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx3","plot_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx4","biosum_plot_id");
			}

			public string CreateDisconnectedRoadTravelTimeOfZeroTableSQL(string p_strTableName)
			{
				return "CREATE TABLE disconnected_road_travel_time_of_zero (" +
					"traveltime_id LONG," + 
					"psite_id INTEGER," + 
					"biosum_plot_id CHAR(24)," + 
					"plot_id LONG)";
			}
			public void CreateProcessingSiteTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateProcessingSiteTableSQL(p_strTableName));
				CreateProcessingSiteTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateProcessingSiteTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","psite_id");
			}
			public string CreateProcessingSiteTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"psite_id INTEGER," + 
					"name CHAR(30)," + 
					"trancd BYTE," + 
					"trancd_def CHAR(20)," + 
					"biocd BYTE," + 
					"biocd_def CHAR(15)," + 
					"exists_yn CHAR(1) DEFAULT 'N'," + 
					"keeplyr_yn CHAR(1) DEFAULT 'N'," + 
					"lat DOUBLE," + 
					"lon DOUBLE)";
			}
			public void CreateTravelTimeTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateTravelTimeTableSQL(p_strTableName));
				CreateTravelTimeTableIndexes(p_oAdo,p_oConn,p_strTableName);


			}
			public void CreateTravelTimeTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","traveltime_id");
				p_oAdo.AddAutoNumber(p_oConn,p_strTableName,"traveltime_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","psite_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx3","plot_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx4","collector_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx5","railhead_id");
			}

			public string CreateTravelTimeTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"traveltime_id LONG," + 
					"psite_id INTEGER," + 
					"biosum_plot_id CHAR(24)," + 
					"plot_id LONG," + 
					"collector_id LONG," + 
					"railhead_id LONG," + 
					"travel_mode BYTE," + 
					"travel_time DOUBLE)";
			}
			public void CreateTravelTimeOfZeroTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateTravelTimeOfZeroTableSQL(p_strTableName));
				CreateTravelTimeOfZeroTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateTravelTimeOfZeroTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","traveltime_id");
				p_oAdo.AddAutoNumber(p_oConn,p_strTableName,"traveltime_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","psite_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx2","biosum_plot_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx3","plot_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx4","collector_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx5","railhead_id");
			}
			public string CreateTravelTimeOfZeroTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"traveltime_id LONG," + 
					"psite_id INTEGER," + 
					"biosum_plot_id CHAR(24)," +
					"plot_id LONG," + 
					"collector_id LONG," + 
					"railhead_id LONG," + 
					"travel_mode BYTE," + 
					"travel_time DOUBLE)";
			}
		}
		public class Audit
		{
			public string DefaultPlotCondAuditTableDbFile {get {return @"db\audit.mdb";}}
			public string DefaultPlotCondAuditTableName {get {return "plot_audit";}}
			public string DefaultPlotCondRxAuditTableDbFile {get {return @"db\audit.mdb";}}
			public string DefaultPlotCondRxAuditTableName {get {return "plot_cond_rx_audit";}}

			public Audit()
			{
			}
			public void CreatePlotCondAuditTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreatePlotCondAuditTableSQL(p_strTableName));
				CreatePlotCondAuditTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreatePlotCondAuditTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id");
			}
			public string CreatePlotCondAuditTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"biosum_cond_id CHAR(25)," + 
					"fvs_prepost_variables_yn CHAR(1)," + 
					"gis_travel_times_yn CHAR(1)," + 
					"processor_tree_vol_val_yn CHAR(1), " + 
					"frcs_harvest_costs_yn CHAR(1))";
			}
			public void CreatePlotCondRxAuditTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreatePlotCondRxAuditTableSQL(p_strTableName));
				CreatePlotCondRxAuditTableIndexes(p_oAdo,p_oConn,p_strTableName);

			}
			public void CreatePlotCondRxAuditTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn, p_strTableName,p_strTableName + "_pk","biosum_cond_id,rxpackage,rx,rxcycle");
			}
			public string CreatePlotCondRxAuditTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"biosum_cond_id CHAR(25)," + 
                    "rxpackage CHAR(3)," + 
                    "rx CHAR(3)," + 
					"rxcycle CHAR(1)," + 
					"fvs_prepost_variables_yn CHAR(1)," + 
					"processor_tree_vol_val_yn CHAR(1)," + 
					"frcs_harvest_costs_yn CHAR(1))";
			}

		}
		public class FIAPlot
		{
			public FIAPlot()
			{
			}
			public string DefaultPlotTableDbFile {get {return @"db\master.mdb";}}
			public string DefaultPlotTableName {get {return "plot";}}
			public string DefaultConditionTableDbFile {get {return @"db\master.mdb";}}
			public string DefaultConditionTableName {get {return "cond";}}
			public string DefaultTreeTableDbFile {get {return @"db\master.mdb";}}
			public string DefaultTreeTableName {get {return "tree";}}
			public string DefaultTreeRegionalBiomassTableDbFile {get {return @"db\master.mdb";}}
			public string DefaultTreeRegionalBiomassTableName {get {return "tree_regional_biomass";}}
			public string DefaultPopEstnUnitTableDbFile {get {return @"db\master.mdb";}}
			public string DefaultPopEstnUnitTableName {get {return "pop_estn_unit";}}
			public string DefaultPopEvalTableDbFile {get {return @"db\master.mdb";}}
			public string DefaultPopEvalTableName {get {return "pop_eval";}}
			public string DefaultPopPlotStratumAssgnTableDbFile {get {return @"db\master.mdb";}}
			public string DefaultPopPlotStratumAssgnTableName {get {return "pop_plot_stratum_assgn";}}
			public string DefaultPopStratumTableDbFile {get {return @"db\master.mdb";}}
			public string DefaultPopStratumTableName {get {return "pop_stratum";}}
            public string DefaultBiosumPopStratumAdjustmentFactorsTableName { get { return "biosum_pop_stratum_adjustment_factors"; } }
            public string DefaultBiosumPopStratumAdjustmentFactorsTableDbFile { get { return @"db\master.mdb"; } }
			public string DefaultSiteTreeTableDbFile {get {return @"db\master.mdb";}}
			public string DefaultSiteTreeTableName {get {return "sitetree";}}




			public void CreatePlotTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreatePlotTableSQL(p_strTableName));
				CreatePlotTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreatePlotTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_plot_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","num_cond");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx2","biosum_status_cd");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx3","gis_status_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx4","idb_plot_id");
			}
			public string CreatePlotTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"biosum_plot_id CHAR(24)," + 
					"statecd INTEGER," + 
					"invyr INTEGER," + 
					"unitcd INTEGER," + 
					"countycd INTEGER," + 
					p_strTableName + " LONG," + 
					"measyear INTEGER," + 
					"measmon INTEGER," + 
					"measday INTEGER," + 
					"elev INTEGER," + 
					"fvs_variant CHAR(2)," + 
					"half_state CHAR(10)," + 
					"subplot_count_plot BYTE," + 
					"gis_yard_dist DOUBLE," + 
					"gis_protected_area_yn CHAR(1)," + 
					"gis_roadless_yn CHAR(1)," + 
					"all_cond_not_accessible_yn CHAR(1)," + 
					p_strTableName + "_accessible_yn CHAR(1)," + 
					"num_cond BYTE," + 
					"one_cond_yn CHAR(1)," + 
					"lat DOUBLE," + 
					"lon DOUBLE," + 
                    "macro_breakpoint_dia INTEGER," + 
					"merch_haul_cost_id LONG," + 
					"merch_haul_cost_psite INTEGER," + 
					"merch_haul_cpa_pt DOUBLE," + 
					"chip_haul_cost_id LONG," + 
					"chip_haul_cost_psite INTEGER," + 
					"chip_haul_cpa_pt DOUBLE," + 
					"biosum_status_cd BYTE," + 
					"gis_status_id INTEGER," + 
					"idb_plot_id LONG," + 
					"cn CHAR(34))";

			}
			
			public void CreateConditionTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateConditionTableSQL(p_strTableName));
				CreateConditionTableIndexes(p_oAdo,p_oConn,p_strTableName);
				
			}
			public void CreateConditionTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","biosum_plot_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx2","condid");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx3","fvs_filename");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx4","idb_cond_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx5","idb_plot_id");
			}
			public string CreateConditionTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"biosum_cond_id CHAR(25)," + 
					"biosum_plot_id CHAR(24)," + 
					"invyr INTEGER," + 
					"condid BYTE," + 
					"condprop DOUBLE," + 
					"landclcd BYTE," + 
					"fortypcd INTEGER," + 
					"ground_land_class_pnw CHAR(3)," + 
					"owncd INTEGER," + 
					"owngrpcd INTEGER," + 
					"reservcd BYTE," + 
					"siteclcd BYTE," + 
					"sibase INTEGER," + 
					"sicond INTEGER," + 
					"sisp INTEGER," + 
					"slope INTEGER," + 
					"aspect INTEGER," + 
					"stdage INTEGER," + 
					"stdszcd BYTE," + 
					"habtypcd1 CHAR(10)," + 
					"adforcd INTEGER," + 
					"qmd_tot_cm SINGLE," + 
					"hwd_qmd_tot_cm SINGLE," + 
					"swd_qmd_tot_cm SINGLE," + 
					"acres DOUBLE," + 
					"unitcd LONG," + 
					"vol_loc_grp CHAR(10)," + 
					"tpacurr DOUBLE," + 
					"hwd_tpacurr DOUBLE," + 
					"swd_tpacurr DOUBLE," + 
					"ba_ft2_ac DOUBLE," + 
					"hwd_ba_ft2_ac DOUBLE," + 
					"swd_ba_ft2_ac DOUBLE," + 
					"vol_ac_grs_stem_ttl_ft3 DOUBLE," + 
					"hwd_vol_ac_grs_stem_ttl_ft3 DOUBLE," + 
					"swd_vol_ac_grs_stem_ttl_ft3 DOUBLE," + 
					"vol_ac_grs_ft3 DOUBLE," + 
					"hwd_vol_ac_grs_ft3 DOUBLE," + 
					"swd_vol_ac_grs_ft3 DOUBLE," + 
					"volcsgrs DOUBLE," + 
					"hwd_volcsgrs DOUBLE," + 
					"swd_volcsgrs DOUBLE," + 
					"gsstkcd DOUBLE," + 
					"alstkcd DOUBLE," + 
					"sdi DOUBLE," + 
					"ccf DOUBLE," + 
					"topht DOUBLE," + 
					"condprop_unadj DOUBLE," + 
                    "micrprop_unadj DOUBLE," + 
                    "subpprop_unadj DOUBLE," + 
                    "macrprop_unadj DOUBLE," + 
					"harvest_technique CHAR(30)," + 
					"cond_too_far_steep_yn CHAR(1) DEFAULT 'N'," + 
					"cond_accessible_yn CHAR(1) DEFAULT 'Y'," + 
					"fvs_filename CHAR(12)," +
					"idb_cond_id LONG," + 
					"idb_plot_id LONG," + 
					"cn CHAR(34)," + 
					"biosum_status_cd BYTE)";

			}
			public void CreateTreeTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateTreeTableSQL(p_strTableName));
				CreateTreeTableIndexes(p_oAdo,p_oConn,p_strTableName);
				
			}
			public void CreateTreeTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","biosum_cond_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx2","subp");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx3","tree");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx4","condid");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx5","idb_dmg_agent3_cd");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx6","idb_severity3_cd");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx7","fvs_tree_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx8","idb_alltree_id");
			}
			public string CreateTreeTableSQL(string p_strTableName)
			{
				return "CREATE TABLE tree (" +
					"biosum_cond_id CHAR(25)," + 
					"invyr INTEGER," + 
					"statecd INTEGER," + 
					"unitcd INTEGER," + 
					"countycd INTEGER," + 
					"subp BYTE," + 
					"tree INTEGER," + 
					"condid INTEGER," + 
					"statuscd BYTE," + 
					"spcd INTEGER," + 
					"spgrpcd INTEGER," + 
					"dia SINGLE," + 
					"diahtcd BYTE," + 
					"ht DOUBLE," + 
					"htcd BYTE," + 
					"actualht DOUBLE," + 
					"formcl BYTE," + 
					"treeclcd BYTE," + 
					"cr DOUBLE," + 
					"cclcd BYTE DEFAULT 0," + 
					"cull DOUBLE," + 
					"roughcull DOUBLE," + 
					"decaycd BYTE," + 
					"stocking DOUBLE," + 
					"tpacurr DOUBLE," + 
					"wdldstem INTEGER," + 
					"volcfnet DOUBLE," + 
					"volcfgrs DOUBLE," + 
					"volcsnet DOUBLE," + 
					"volcsgrs DOUBLE," + 
					"volbfnet DOUBLE," + 
					"volbfgrs DOUBLE," + 
                    "voltsgrs DOUBLE," + 
					"drybiot DOUBLE," + 
					"drybiom DOUBLE," + 
					"bhage INTEGER," + 
					"cullbf DOUBLE," + 
					"cullsf DOUBLE," + 
					"totage DOUBLE," + 
					"mist_cl_cd INTEGER," + 
					"agentcd INTEGER," + 
					"damtyp1 INTEGER," + 
					"damsev1 INTEGER," + 
					"damtyp2 INTEGER," + 
					"damsev2 INTEGER," + 
					"tpa_unadj DOUBLE," + 
					"idb_dmg_agent3_cd	INTEGER," + 
					"idb_severity3_cd INTEGER," + 
					"fvs_dmg_ag1 CHAR(2)," + 
					"fvs_dmg_sv1 CHAR(2)," + 
					"fvs_dmg_ag2 CHAR(2)," + 
					"fvs_dmg_sv2 CHAR(2)," + 
					"fvs_dmg_ag3 CHAR(2)," + 
					"fvs_dmg_sv3 CHAR(2)," + 
                    "inc10yr INTEGER," +
                    "condprop_specific DOUBLE," + 
					"fvs_tree_id CHAR(10)," + 
					"idb_alltree_id LONG," + 
					"cn CHAR(34)," + 
					"biosum_status_cd BYTE)";

			}
			public void CreateSiteTreeTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateSiteTreeTableSQL(p_strTableName));
				CreateSiteTreeTableIndexes(p_oAdo,p_oConn,p_strTableName);

			}
			public void CreateSiteTreeTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","condid");
			}
			public string CreateSiteTreeTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"biosum_plot_id CHAR(24)," + 
					"invyr INTEGER," + 
					"condid BYTE," + 
					"tree INTEGER," + 
					"spcd INTEGER," + 
					"dia SINGLE," + 
					"ht DOUBLE," + 
					"agedia INTEGER," + 
					"spgrpcd INTEGER," + 
					"sitree INTEGER," + 
					"sibase INTEGER," + 
					"subp BYTE," + 
					"method INTEGER," + 
					"sitree_est INTEGER," + 
					"validcd BYTE," + 
					"condlist INTEGER," + 
					"biosum_status_cd BYTE)";

			}
			public void CreateTreeRegionalBiomassTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateTreeRegionalBiomassTableSQL(p_strTableName));
				CreateTreeRegionalBiomassTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateTreeRegionalBiomassTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","tre_cn");
			}
			public string CreateTreeRegionalBiomassTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"tre_cn CHAR(34)," + 
					"statecd INTEGER," + 
					"regional_drybiot DOUBLE," + 
					"regional_drybiom DOUBLE," + 
					"biosum_status_cd BYTE)";
			}
			public void CreatePopEstnUnitTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreatePopEstnUnitTableSQL(p_strTableName));
				CreatePopEstnUnitTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreatePopEstnUnitTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","evalid");
			}
			public string CreatePopEstnUnitTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"cn CHAR(34)," + 
					"eval_cn CHAR(34)," + 
					"rscd DECIMAL (2,0)," + 
					"evalid DECIMAL (6,0)," +
					"estn_unit_descr CHAR(255)," + 
					"statecd DECIMAL (4,0)," + 
					"arealand_eu DECIMAL (12,2)," + 
					"areatot_eu DECIMAL (12,2)," + 
					"area_used DECIMAL (12,2)," + 
					"area_source CHAR(50)," + 
					"p1pntcnt_eu DECIMAL (12,0)," + 
					"p1source CHAR(30)," + 
					"biosum_status_cd BYTE)";

			}
			public void CreatePopEvalTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreatePopEvalTableSQL(p_strTableName));
                CreatePopEvalTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreatePopEvalTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","evalid");
			}
			public string CreatePopEvalTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"cn CHAR(34)," + 
					"rscd DECIMAL (2,0)," + 
					"evalid DECIMAL (6,0)," +
					"eval_descr CHAR(255)," + 
					"statecd DECIMAL (4,0)," + 
					"location_nm CHAR(255)," + 
					"report_year_nm CHAR(255)," + 
					"notes MEMO," + 
					"start_invyr DECIMAL (4,0)," + 
					"end_invyr DECIMAL (4,0)," + 
					"p1source CHAR(30)," + 
					"biosum_status_cd BYTE)";
			}
			public void CreatePopPlotStratumAssgnTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreatePopPlotStratumAssgnTableSQL(p_strTableName));
				CreatePopPlotStratumAssgnTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreatePopPlotStratumAssgnTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","evalid");
			}
			public string CreatePopPlotStratumAssgnTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"cn CHAR(34)," + 
					"stratum_cn CHAR(34)," + 
					"plt_cn CHAR(34)," + 
					"statecd DECIMAL (4,0)," + 
					"invyr INTEGER," + 
					"unitcd INTEGER," + 
					"countycd INTEGER," + 
					"plot LONG," + 
					"rscd DECIMAL (2,0)," + 
					"evalid DECIMAL (6,0)," +
					"estn_unit LONG," + 
					"stratumcd LONG," + 
					"biosum_status_cd BYTE)";
			}
			public void CreatePopStratumTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreatePopStratumTableSQL(p_strTableName));
				CreatePopStratumTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreatePopStratumTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","evalid");
			}
			public string CreatePopStratumTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"cn CHAR(34)," + 
					"estn_unit_cn CHAR(34)," + 
					"rscd DECIMAL (2,0)," + 
					"evalid DECIMAL (6,0)," +
					"estn_unit LONG," + 
					"stratumcd LONG," + 
					"statum_desc CHAR(255)," + 
					"statecd DECIMAL (4,0)," + 
					"p1pointcnt DECIMAL (12,0)," + 
					"p2pointcnt DECIMAL (12,0)," + 
					"expns DOUBLE," +
					"adj_factor_macr DECIMAL (5,4)," + 
					"adj_factor_subp DECIMAL (5,4)," + 
					"adj_factor_micr DECIMAL (5,4)," + 
					"biosum_status_cd BYTE)";
			}
            public void CreateBiosumPopStratumAdjustmentFactorsTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateBiosumPopStratumAdjustmentFactorsTableSQL(p_strTableName));
                CreateBiosumPopStratumAdjustmentFactorsTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateBiosumPopStratumAdjustmentFactorsTableIndexes(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "rscd,evalid");
            }
            public string CreateBiosumPopStratumAdjustmentFactorsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "stratum_cn CHAR(34)," +
                    "rscd DECIMAL (2,0)," +
                    "evalid DECIMAL (6,0)," +
                    "eval_descr CHAR(255)," + 
                    "estn_unit LONG," +
                    "estn_unit_descr CHAR(255)," + 
                    "stratumcd LONG," +
                    "p2pointcnt_man DECIMAL (12,0)," +
                    "double_sampling INTEGER," +
                    "stratum_area DOUBLE," + 
                    "expns DOUBLE," +
                    "pmh_macr DECIMAL (5,4)," +
                    "pmh_sub DECIMAL (5,4)," +
                    "pmh_micr DECIMAL (5,4)," +
                    "pmh_cond DECIMAL (5,4)," +
                    "adj_factor_macr DECIMAL (5,4)," +
                    "adj_factor_subp DECIMAL (5,4)," +
                    "adj_factor_micr DECIMAL (5,4)," +
                    "biosum_status_cd BYTE)";
            }
		}
        public class Processor
        {
            public Processor()
            {
            }
            static public string DefaultHarvestCostsTableDbFile { get { return @"db\master.mdb"; } }
            static public string DefaultHarvestCostsTableName { get { return "harvest_costs"; } }
            static public string DefaultAdditionalHarvestCostsTableName { get { return "additional_harvest_costs"; } }
            static public string DefaultAdditionalHarvestCostsTableDbFile { get { return @"db\master.mdb"; } }
            static public string DefaultTreeVolValSpeciesDiamGroupsDbFile { get { return @"db\master.mdb"; } }
            static public string DefaultTreeVolValSpeciesDiamGroupsTableName { get { return "tree_vol_val_by_species_diam_groups"; } }
            static public string DefaultHarvestMethodTableDbFile { get { return @"db\master.mdb"; } }
            static public string DefaultHarvestMethodTableName { get { return @"harvest_method"; } }

            public void CreateHarvestCostsTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, Tables.Processor.CreateHarvestCostsTableSQL(p_strTableName));
                CreateHarvestCostsTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateHarvestCostsTableIndexes(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "biosum_cond_id,rxpackage,rx,rxcycle");
            }
            static public string CreateHarvestCostsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "complete_cpa DOUBLE DEFAULT 0," +
                    "harvest_cpa DOUBLE ," +
                    "chip_cpa DOUBLE ," +
                    "assumed_movein_cpa DOUBLE ," +
                    "ideal_complete_cpa DOUBLE DEFAULT 0," +
                    "ideal_harvest_cpa DOUBLE ," +
                    "ideal_chip_cpa DOUBLE ," +
                    "ideal_assumed_movein_cpa DOUBLE ," +
                    "harvest_cpa_warning_msg CHAR(240)," +
				//						"water_barring_roads_cpa DOUBLE," +
                //    "brush_cutting_cpa DOUBLE," +                     
                    "place_holder CHAR(1) DEFAULT 'N'," +
                    "override_YN CHAR(1) DEFAULT 'N'," +
                    "DateTimeCreated CHAR(22))";

            }
            public void CreateAdditionalHarvestCostsTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, Tables.Processor.CreateAdditionalHarvestCostsTableSQL(p_strTableName));
                CreateAdditionalHarvestCostsTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateAdditionalHarvestCostsTableIndexes(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "biosum_cond_id,rx");
            }
            static public string CreateAdditionalHarvestCostsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rx CHAR(3))";
                   // "water_barring_roads_cpa DOUBLE," +
                   // "brush_cutting_cpa DOUBLE)";

            }
            public void CreateHarvestTechniqueTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateHarvestTechniqueTableSQL(p_strTableName));
                CreateHarvestTechniqueTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateHarvestTechniqueTableIndexes(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "biosum_cond_id,rx");
            }
            public string CreateHarvestTechniqueTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "harvest_technique CHAR(30)," +
                    "harvest_technique_40 CHAR(30))";

            }

            public void CreateTreeVolValSpeciesDiamGroupsTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, Tables.Processor.CreateTreeVolValSpeciesDiamGroupsTableSQL(p_strTableName));
                CreateTreeVolValSpeciesDiamGroupsTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateTreeVolValSpeciesDiamGroupsTableIndexes(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddAutoNumber(p_oConn, p_strTableName, "id");
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "id");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "biosum_cond_id,rxpackage,rx,rxcycle");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "rx");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx3", "species_group");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx4", "diam_group");
            }

            static public string CreateTreeVolValSpeciesDiamGroupsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "id LONG," +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "species_group INTEGER," +
                    "diam_group INTEGER DEFAULT 0," +
                    "biosum_harvest_method_category INTEGER DEFAULT 0," +
                    "chip_vol_cf DOUBLE," +
                    "chip_wt_gt DOUBLE," +
                    "chip_val_dpa DOUBLE," +
                    "chip_mkt_val_pgt DOUBLE DEFAULT 0," +
                    "merch_vol_cf DOUBLE," +
                    "merch_wt_gt DOUBLE," +
                    "merch_val_dpa DOUBLE," +
                    "merch_to_chipbin_YN CHAR(1) DEFAULT 'N'," +
                    "bc_vol_cf DOUBLE," +
                    "bc_wt_gt DOUBLE," +
                    "stand_residue_wt_gt DOUBLE," +
                    "place_holder CHAR(1) DEFAULT 'N'," +
                    "DateTimeCreated CHAR(22))";
            }
            public void CreateTreeVolValSpeciesDiamGroupsWorkTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, Tables.Processor.CreateTreeVolValSpeciesDiamGroupsWorkTableSQL(p_strTableName));
                CreateTreeVolValSpeciesDiamGroupsWorkTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateTreeVolValSpeciesDiamGroupsWorkTableIndexes(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "biosum_cond_id,rxpackage,rx,rxcycle");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "rx");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx3", "species_group");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx4", "diam_group");
            }

            static public string CreateTreeVolValSpeciesDiamGroupsWorkTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "species_group INTEGER," +
                    "diam_group INTEGER DEFAULT 0," +
                    "chip_vol_cf DOUBLE," +
                    "chip_wt_gt DOUBLE," +
                    "chip_val_dpa DOUBLE," +
                    "chip_mkt_val_pgt DOUBLE DEFAULT 0," +
                    "merch_vol_cf DOUBLE," +
                    "merch_wt_gt DOUBLE," +
                    "merch_val_dpa DOUBLE," +
                    "merch_to_chipbin_YN CHAR(1) DEFAULT 'N'," +
                    "bc_wt_gt DOUBLE," +
                    "bc_vol_cf DOUBLE," +
                    "DateTimeCreated CHAR(22))";
            }

            public void CreateOpcostInputTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, Tables.Processor.CreateOpcostInputTableSQL(p_strTableName));
                // No indexes currently on OpCost input table
            }


            static public string CreateOpcostInputTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " " +
                             "(Stand text (255)," +
                             " [Percent Slope] short," +
                             " [One-way Yarding Distance] DOUBLE," +
                             " YearCostCalc long," +
                             " [Project Elevation] short," +
                             " [Harvesting System] text (50)," +
                             " [Chip tree per acre] single," +
                             " [Residue fraction for chip trees] single," +
                            " [Chip trees average volume(ft3)] single," +
                            " [CHIPS Average Density (lbs/ft3)] single," +
                            " [CHIPS Hwd Proportion] single," +
                            " [Small log trees per acre] single," +
                            " [Small log trees residue fraction] single," +
                            " [Small log trees average volume(ft3)] single," +
                            " [Small log trees average density(lbs/ft3)] single," +
                            " [Small log trees hardwood proportion] single," +
                            " [Large log trees per acre] single," +
                            " [Large log trees residue fraction] single," +
                            " [Large log trees average vol(ft3)] single," +
                            " [Large log trees average density(lbs/ft3)] single," +
                            " [Large log trees hardwood proportion] single," +
                            " BrushCutTPA single," +
                            " BrushCutAvgVol single," +
                            " RxPackage_Rx_RxCycle text (255))";
            }


            public void CreateNewOpcostInputTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, Tables.Processor.CreateNewOpcostInputTableSQL(p_strTableName));
                // No indexes currently on OpCost input table
            }

            static public string CreateNewOpcostInputTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " " +
                             "(Stand text (255)," +
                             " [Percent Slope] short," +
                             " [One-way Yarding Distance] DOUBLE," +
                             " YearCostCalc long," +
                             " [Project Elevation] short," +
                             " [Harvesting System] text (50)," +
                             " [Chip tree per acre] single," +
                             " [Chip trees MerchAsPctOfTotal] single," +
                             " [Chip trees average volume(ft3)] single," +
                            " [CHIPS Average Density (lbs/ft3)] single," +
                            " [CHIPS Hwd Percent] single," +
                            " [Small log trees per acre] single," +
                            " [Small log trees MerchAsPctOfTotal] single," +
                            " [Small log trees ChipPct_Cat1_3] single," +
                            " [Small log trees ChipPct_Cat2_4] single," +
                            " [Small log trees ChipPct_Cat5] single," +
                            " [Small log trees average volume(ft3)] single," +
                            " [Small log trees average density(lbs/ft3)] single," +
                            " [Small log trees hardwood percent] single," +
                            " [Large log trees per acre] single," +
                            " [Large log trees MerchAsPctOfTotal] single," +
                            " [Large log trees ChipPct_Cat1_3_4] single," +
                            " [Large log trees ChipPct_Cat2] single," +
                            " [Large log trees ChipPct_Cat5] single," +
                            " [Large log trees average vol(ft3)] single," +
                            " [Large log trees average density(lbs/ft3)] single," +
                            " [Large log trees hardwood percent] single," +
                            " BrushCutTPA single," +
                            " BrushCutAvgVol single," +
                            " RxPackage_Rx_RxCycle text (255)," +
                            " RxCycle text (255), " +
                            " biosum_cond_id text (255)," +
                            " RxPackage text (255)," +
                            " Rx text (255)," +
                            " Move_In_Hours single," +
                            " Harvest_area_assumed_acres single," +
                            " [Unadjusted One-way Yarding distance] double," +
                            " [Unadjusted Small log trees per acre] single," +
                            " [Unadjusted Small log trees average volume (ft3)] single," +
                            " [Unadjusted Large log trees per acre] single," +
                            " [Unadjusted Large log trees average vol(ft3)] single" +
                            " )";
            }

            public void CreateTreeReconcilationTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, Tables.Processor.CreateTreeReconcilationTableSQL(p_strTableName));
                // No indexes currently on OpCost input table
            }

            static public string CreateTreeReconcilationTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName +
                       " (cn text (34)," +
                       " tree long," +
                       " biosum_cond_id text (25)," +
                       " biosum_plot_id text (24)," +
                       " spcd INTEGER," +
                       " merchWtGt double," +
                       " nonMerchWtGt double," +
                       " drybiom double," +
                       " drybiot double," +
                       " volCfNet double," +
                       " volCfGrs double," +
                       " volTsGrs double," +
                       " odWgt double," +
                       " dryToGreen double, " +
                       " tpa double, " +
                       " dbh double, " +
                       " isSapling bit, " +
                       " isWoodland bit, " +
                       " isCull bit, " +
                       " species_group integer, " +
                       " diam_group integer, " +
                       " merch_value double, " +
                       " opcost_type text (5), " +
                       " biosum_category integer)";
            }
        }

		public class Scenario
		{
			private string strSQL = "";
			public static string DefaultScenarioTableName {get {return "scenario";}}
			public static string DefaultScenarioDatasourceTableName {get {return "scenario_datasource";}}

			public Scenario()
			{
			}
			public void CreateScenarioTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateScenarioTableSQL(p_strTableName));
				CreateScenarioTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateScenarioTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","scenario_id");
			}
			public static string CreateScenarioTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," + 
					"description MEMO," + 
					"path CHAR(254)," + 
					"file CHAR(50)," + 
					"notes MEMO)";

				
			}
			
			public void CreateScenarioDatasourceTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateScenarioDatasourceTableSQL(p_strTableName));
				CreateScenarioDatasourceTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateScenarioDatasourceTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx","scenario_id");
			}
			public static string CreateScenarioDatasourceTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," + 
					"table_type CHAR(60)," + 
					"path CHAR(254)," + 
					"file CHAR(50)," + 
					"table_name CHAR(50))";

				
			}

		}
		public class ProcessorScenarioRun
		{
			private string strSQL = "";
			static public string DefaultHarvestCostsTableDbFile {get {return @"db\scenario_results.mdb";}}
			static public string DefaultHarvestCostsTableName {get {return "harvest_costs";}}
			static public string DefaultTreeVolValSpeciesDiamGroupsDbFile {get {return @"db\scenario_results.mdb";}}
			static public string DefaultTreeVolValSpeciesDiamGroupsTableName {get {return "tree_vol_val_by_species_diam_groups";}}
            static public string DefaultDiametersLowSlopeTableName {get { return "DiametersLowSlope"; } }
            static public string DefaultDiametersSteepSlopeTableName { get { return "DiametersSteepSlope"; } }
            static public string DefaultTreeDataInLowSlopeTableName { get { return "TreeDataInLowSlope"; } }
            static public string DefaultTreeDataInSteepSlopeTableName { get { return "TreeDataInSteepSlope"; } }
            static public string DefaultTreeBinLowSlopeTableName { get { return "BinsLowSlope"; } }
            static public string DefaultTreeHwdBinLowSlopeTableName { get { return "Hwd_BinsLowSlope"; } }
            static public string DefaultTreeBinSteepSlopeTableName { get { return "BinsSteepSlope"; } }
            static public string DefaultTreeHwdBinSteepSlopeTableName { get { return "Hwd_BinsSteepSlope"; } }
            static public string DefaultFiaTreeSpeciesRefTableName { get { return "FIA_TREE_SPECIES_REF"; } }

			public ProcessorScenarioRun()
			{
			}
			//
			//HARVEST COSTS TABLE
			//
			static public string CreateHarvestCostsTableSQL(string p_strTableName)
			{
				return   Tables.Processor.CreateHarvestCostsTableSQL(p_strTableName);
			}
			//
			//TREE VOLUMES AND VALUES BY SPECIES GROUPS AND DIAMETER GROUPS
			//
			static public string CreateTreeVolValSpeciesDiamGroupsTableSQL(string p_strTableName)
			{
				return Tables.Processor.CreateTreeVolValSpeciesDiamGroupsTableSQL(p_strTableName);
			}
            //
            //DIAMETERS TABLE
            //
            public void CreateDiametersTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateDiametersTableSQL(p_strTableName));
                CreateDiametersTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateDiametersTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "dbh,diam_group,species_group");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "diam_group");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "species_group");
            }
            static public string CreateDiametersTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " " + 
                       "(DBH Single, class text (10), Util_Logs_Bk text (1), " + 
                        "Util_Chips_Bk text (1),Res_Stmp_Bk text (1)," + 
                        "Res_Lnd_Bk text (1), Util_Logs_Br text (1)," + 
                        "Util_Chips_Br text (1)," + 
                        "Res_Stmp_Br text (1), Res_Lnd_Br text (1), " +
                        "Util_Logs_Bl text (1), Util_Chips_Bl text (1)," + 
                        "Res_Stmp_Bl text (1), Res_Lnd_Bl text (1)," + 
                        "NonUtil_Logs_Bl text (1)," + 
                        "NonUtil_Chips_Bl text (1)," + 
                        "diam_group short, species_group INTEGER)";

            }
            //
            //TREE DATA_IN
            //
            public void CreateTreeDataInTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateTreeDataInTableSQL(p_strTableName));
                CreateTreeDataInTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateTreeDataInTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "biosum_cond_id");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "spcd");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx3", "species_group");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx4", "fvs_tree_id");
            }
            static public string CreateTreeDataInTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " " +
                       "(fvs_tree_id CHAR(10), " +
                        "biosum_cond_id CHAR(25)," +
                        "rxpackage CHAR(3)," + 
                        "rx CHAR(3)," + 
                        "rxcycle CHAR(1)," +
                        "spcd INTEGER," + 
                        "dia SINGLE," + 
                        "tpa DOUBLE," + 
                        "merch_vol_cf DOUBLE," + 
                        "merch_wt_gt DOUBLE," + 
                        "chip_vol_cf DOUBLE," + 
                        "chip_wt_gt DOUBLE," + 
                        "residue_fraction DOUBLE," + 
                        "slope INTEGER," + 
                        "species_group INTEGER," + 
                        "diam_group INTEGER," + 
                        "merch_to_chipbin_YN CHAR(1) DEFAULT 'N')";

            }
            //
            //TREE BIN
            //
            public void CreateTreeBinTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateTreeBinTableSQL(p_strTableName));
                CreateTreeBinTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateTreeBinTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "fvs_tree_id");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "biosum_cond_id");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx3", "species");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx4", "species_group");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx5", "diam_group");
            }
            static public string CreateTreeBinTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " " +
                         "(FVS_TREE_ID TEXT (10)," + 
                          "BIOSUM_COND_ID text (25)," + 
                          "RXPACKAGE text (3)," + 
                          "RX text (3)," +
                          "RXCYCLE text (1)," +  
                          "species short," +  
  						  "species_group short," + 
  						  "DBH single," + 
  						  "diam_group short," +
						  "BC_Util_Logs_count short," + 
						  "BC_Util_Logs_TPA single," + 
						  "BC_Util_Logs_merch_vol single," +
						  "BC_Util_Logs_merch_wt single," + 
						  "BC_Util_Logs_biomass_vol single," + 
						  "BC_Util_Logs_biomass_wt single," + 
						  "BC_Util_Chips_count integer," +
						  "BC_Util_Chips_TPA single," +
						  "BC_Util_Chips_merch_wt single," + 
						  "BC_Util_Chips_merch_vol single," +
						  "BC_Util_Chips_biomass_wt single," +
						  "BC_Util_Chips_biomass_vol single," +
                          "BC_NonUtil_Logs_count short," +
                          "BC_NonUtil_Logs_TPA single," +
                          "BC_NonUtil_Logs_merch_vol single," +
                          "BC_NonUtil_Logs_merch_wt single," +
                          "BC_NonUtil_Logs_biomass_vol single," +
                          "BC_NonUtil_Logs_biomass_wt single," +
                          "BC_NonUtil_Chips_count integer," +
                          "BC_NonUtil_Chips_TPA single," +
                          "BC_NonUtil_Chips_merch_wt single," +
                          "BC_NonUtil_Chips_merch_vol single," +
                          "BC_NonUtil_Chips_biomass_wt single," +
                          "BC_NonUtil_Chips_biomass_vol single," +
						  "CHIPS_Util_Logs_count integer," +
						  "CHIPS_Util_Logs_TPA single," +
						  "CHIPS_Util_Chips_count integer," + 
						  "CHIPS_Util_Chips_TPA single," +
						  "CHIPS_Util_Logs_merch_vol single," +
						  "CHIPS_Util_Logs_merch_wt single," +
						  "CHIPS_Util_logs_biomass_vol single," +
						  "CHIPS_Util_logs_biomass_wt single," +
						  "CHIPS_Util_Chips_merch_vol single," +
						  "CHIPS_Util_Chips_merch_wt single," + 
						  "CHIPS_Util_Chips_biomass_vol single," +
						  "CHIPS_Util_Chips_biomass_wt single," +
                          "CHIPS_NonUtil_Logs_count integer," +
                          "CHIPS_NonUtil_Logs_TPA single," +
                          "CHIPS_NonUtil_Chips_count integer," +
                          "CHIPS_NonUtil_Chips_TPA single," +
                          "CHIPS_NonUtil_Logs_merch_vol single," +
                          "CHIPS_NonUtil_Logs_merch_wt single," +
                          "CHIPS_NonUtil_logs_biomass_vol single," +
                          "CHIPS_NonUtil_logs_biomass_wt single," +
                          "CHIPS_NonUtil_Chips_merch_vol single," +
                          "CHIPS_NonUtil_Chips_merch_wt single," +
                          "CHIPS_NonUtil_Chips_biomass_vol single," +
                          "CHIPS_NonUtil_Chips_biomass_wt single," + 
						  "SMLOGS_Util_Logs_count integer," +
						  "SMLOGS_Util_Logs_TPA single," + 
						  "SMLOGS_Util_Chips_count integer," +
						  "SMLOGS_Util_Chips_TPA single," + 
						  "SMLOGS_Util_Logs_merch_vol single," + 
						  "SMLOGS_Util_Logs_merch_wt single," +
						  "SMLOGS_Util_logs_biomass_vol single," +
						  "SMLOGS_Util_logs_biomass_wt single," + 
						  "SMLOGS_Util_Chips_merch_vol single," + 
						  "SMLOGS_Util_Chips_merch_wt single," +
						  "SMLOGS_Util_Chips_biomass_vol single," + 
						  "SMLOGS_Util_Chips_biomass_wt single," +
                          "SMLOGS_NonUtil_Logs_count integer," +
                          "SMLOGS_NonUtil_Logs_TPA single," +
                          "SMLOGS_NonUtil_Chips_count integer," +
                          "SMLOGS_NonUtil_Chips_TPA single," +
                          "SMLOGS_NonUtil_Logs_merch_vol single," +
                          "SMLOGS_NonUtil_Logs_merch_wt single," +
                          "SMLOGS_NonUtil_logs_biomass_vol single," +
                          "SMLOGS_NonUtil_logs_biomass_wt single," +
                          "SMLOGS_NonUtil_Chips_merch_vol single," +
                          "SMLOGS_NonUtil_Chips_merch_wt single," +
                          "SMLOGS_NonUtil_Chips_biomass_vol single," +
                          "SMLOGS_NonUtil_Chips_biomass_wt single," + 
						  "LGLOGS_Util_Logs_count integer," +
						  "LGLOGS_Util_Logs_TPA single," +
						  "LGLOGS_Util_Chips_count integer," +
						  "LGLOGS_Util_Chips_TPA single," + 
						  "LGLOGS_Util_Logs_merch_vol single," + 
						  "LGLOGS_Util_Logs_merch_wt single," +
						  "LGLOGS_Util_logs_biomass_vol single," +
						  "LGLOGS_Util_logs_biomass_wt single," +
						  "LGLOGS_Util_Chips_merch_vol single," +
						  "LGLOGS_Util_Chips_merch_wt single," + 
						  "LGLOGS_Util_Chips_biomass_vol single," + 
						  "LGLOGS_Util_Chips_biomass_wt single," + 
                          "LGLOGS_NonUtil_Logs_count integer," +
						  "LGLOGS_NonUtil_Logs_TPA single," +
						  "LGLOGS_NonUtil_Chips_count integer," +
						  "LGLOGS_NonUtil_Chips_TPA single," + 
						  "LGLOGS_NonUtil_Logs_merch_vol single," + 
						  "LGLOGS_NonUtil_Logs_merch_wt single," +
						  "LGLOGS_NonUtil_logs_biomass_vol single," +
						  "LGLOGS_NonUtil_logs_biomass_wt single," +
						  "LGLOGS_NonUtil_Chips_merch_vol single," +
						  "LGLOGS_NonUtil_Chips_merch_wt single," + 
						  "LGLOGS_NonUtil_Chips_biomass_vol single," + 
						  "LGLOGS_NonUtil_Chips_biomass_wt single)";
            }
            //
            //TREE HARDWOOD BIN
            //
            public void CreateTreeHwdBinTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateTreeHwdBinTableSQL(p_strTableName));
                CreateTreeHwdBinTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateTreeHwdBinTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "fvs_tree_id");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "biosum_cond_id");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx3", "species");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx4", "species_group");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx5", "diam_group");
            }
            static public string CreateTreeHwdBinTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " " +
                        "(FVS_TREE_ID text (10), " +
                            "BIOSUM_COND_ID text (25)," +
                            "RXPACKAGE text (3)," +
                            "RX text (3)," +
                            "RXCYCLE text (1)," +
                            "species short," +
                            "species_group short," +
                            "DBH single," +
                            "diam_group short," +
                            "HWD_BC_Util_Logs_count short," +
                            "HWD_BC_Util_Logs_TPA single," +
                            "HWD_BC_Util_Logs_merch_vol single," +
                            "HWD_BC_Util_Logs_merch_wt single," +
                            "HWD_BC_Util_Logs_biomass_vol single," +
                            "HWD_BC_Util_Logs_biomass_wt single," +
                            "HWD_BC_Util_Chips_count integer," +
                            "HWD_BC_Util_Chips_TPA single," +
                            "HWD_BC_Util_Chips_merch_wt single," +
                            "HWD_BC_Util_Chips_merch_vol single," +
                            "HWD_BC_Util_Chips_biomass_wt single," +
                            "HWD_BC_Util_Chips_biomass_vol single," +
                            "HWD_BC_NonUtil_Logs_count short," +
                            "HWD_BC_NonUtil_Logs_TPA single," +
                            "HWD_BC_NonUtil_Logs_merch_vol single," +
                            "HWD_BC_NonUtil_Logs_merch_wt single," +
                            "HWD_BC_NonUtil_Logs_biomass_vol single," +
                            "HWD_BC_NonUtil_Logs_biomass_wt single," +
                            "HWD_BC_NonUtil_Chips_count integer," +
                            "HWD_BC_NonUtil_Chips_TPA single," +
                            "HWD_BC_NonUtil_Chips_merch_wt single," +
                            "HWD_BC_NonUtil_Chips_merch_vol single," +
                            "HWD_BC_NonUtil_Chips_biomass_wt single," +
                            "HWD_BC_NonUtil_Chips_biomass_vol single," +
                            "HWD_CHIPS_Util_Logs_count integer," +
                            "HWD_CHIPS_Util_Logs_TPA single," +
                            "HWD_CHIPS_Util_Chips_count integer," +
                            "HWD_CHIPS_Util_Chips_TPA single," +
                            "HWD_CHIPS_Util_Logs_merch_vol single," +
                            "HWD_CHIPS_Util_Logs_merch_wt single," +
                            "HWD_CHIPS_Util_logs_biomass_vol single," +
                            "HWD_CHIPS_Util_logs_biomass_wt single," +
                            "HWD_CHIPS_Util_Chips_merch_vol single," +
                            "HWD_CHIPS_Util_Chips_merch_wt single," +
                            "HWD_CHIPS_Util_Chips_biomass_vol single," +
                            "HWD_CHIPS_Util_Chips_biomass_wt single," +
                            "HWD_CHIPS_NonUtil_Logs_count integer," +
                            "HWD_CHIPS_NonUtil_Logs_TPA single," +
                            "HWD_CHIPS_NonUtil_Chips_count integer," +
                            "HWD_CHIPS_NonUtil_Chips_TPA single," +
                            "HWD_CHIPS_NonUtil_Logs_merch_vol single," +
                            "HWD_CHIPS_NonUtil_Logs_merch_wt single," +
                            "HWD_CHIPS_NonUtil_logs_biomass_vol single," +
                            "HWD_CHIPS_NonUtil_logs_biomass_wt single," +
                            "HWD_CHIPS_NonUtil_Chips_merch_vol single," +
                            "HWD_CHIPS_NonUtil_Chips_merch_wt single," +
                            "HWD_CHIPS_NonUtil_Chips_biomass_vol single," +
                            "HWD_CHIPS_NonUtil_Chips_biomass_wt single," +
                            "HWD_SMLOGS_Util_Logs_count integer," +
                            "HWD_SMLOGS_Util_Logs_TPA single," +
                            "HWD_SMLOGS_Util_Chips_count integer," +
                            "HWD_SMLOGS_Util_Chips_TPA single," +
                            "HWD_SMLOGS_Util_Logs_merch_vol single," +
                            "HWD_SMLOGS_Util_Logs_merch_wt single," +
                            "HWD_SMLOGS_Util_logs_biomass_vol single," +
                            "HWD_SMLOGS_Util_logs_biomass_wt single," +
                            "HWD_SMLOGS_Util_Chips_merch_vol single," +
                            "HWD_SMLOGS_Util_Chips_merch_wt single," +
                            "HWD_SMLOGS_Util_Chips_biomass_vol single," +
                            "HWD_SMLOGS_Util_Chips_biomass_wt single," +
                            "HWD_SMLOGS_NonUtil_Logs_count integer," +
                            "HWD_SMLOGS_NonUtil_Logs_TPA single," +
                            "HWD_SMLOGS_NonUtil_Chips_count integer," +
                            "HWD_SMLOGS_NonUtil_Chips_TPA single," +
                            "HWD_SMLOGS_NonUtil_Logs_merch_vol single," +
                            "HWD_SMLOGS_NonUtil_Logs_merch_wt single," +
                            "HWD_SMLOGS_NonUtil_logs_biomass_vol single," +
                            "HWD_SMLOGS_NonUtil_logs_biomass_wt single," +
                            "HWD_SMLOGS_NonUtil_Chips_merch_vol single," +
                            "HWD_SMLOGS_NonUtil_Chips_merch_wt single," +
                            "HWD_SMLOGS_NonUtil_Chips_biomass_vol single," +
                            "HWD_SMLOGS_NonUtil_Chips_biomass_wt single," +
                            "HWD_LGLOGS_Util_Logs_count integer," +
                            "HWD_LGLOGS_Util_Logs_TPA single," +
                            "HWD_LGLOGS_Util_Chips_count integer," +
                            "HWD_LGLOGS_Util_Chips_TPA single," +
                            "HWD_LGLOGS_Util_Logs_merch_vol single," +
                            "HWD_LGLOGS_Util_Logs_merch_wt single," +
                            "HWD_LGLOGS_Util_logs_biomass_vol single," +
                            "HWD_LGLOGS_Util_logs_biomass_wt single," +
                            "HWD_LGLOGS_Util_Chips_merch_vol single," +
                            "HWD_LGLOGS_Util_Chips_merch_wt single," + 
                            "HWD_LGLOGS_Util_Chips_biomass_vol single," +
                            "HWD_LGLOGS_Util_Chips_biomass_wt single," +
                            "HWD_LGLOGS_NonUtil_Logs_count integer," +
                            "HWD_LGLOGS_NonUtil_Logs_TPA single," +
                            "HWD_LGLOGS_NonUtil_Chips_count integer," +
                            "HWD_LGLOGS_NonUtil_Chips_TPA single," +
                            "HWD_LGLOGS_NonUtil_Logs_merch_vol single," +
                            "HWD_LGLOGS_NonUtil_Logs_merch_wt single," +
                            "HWD_LGLOGS_NonUtil_logs_biomass_vol single," +
                            "HWD_LGLOGS_NonUtil_logs_biomass_wt single," +
                            "HWD_LGLOGS_NonUtil_Chips_merch_vol single," +
                            "HWD_LGLOGS_NonUtil_Chips_merch_wt single," +
                            "HWD_LGLOGS_NonUtil_Chips_biomass_vol single," +
                            "HWD_LGLOGS_NonUtil_Chips_biomass_wt single)";

            }
            public void CreateOutputTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateOutputTableSQL(p_strTableName));
                CreateOutputTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateOutputTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "biosum_cond_id,rxpackage,rx,rxcycle");
                
            }
            static public string CreateOutputTableSQL(string p_strTableName)
            {
               
               return "CREATE TABLE " + p_strTableName + " " +
                            "(BIOSUM_COND_ID text (25)," +
                            " RXPACKAGE text (3)," +
                            " RX text (3)," +
                            " RXCYCLE text (1)," +
                            " elev short," +
                            " slope short," +
                            " gis_yard_dist DOUBLE," + 
                            " [Harvesting system] text (50)," +
                           " [CHIPS TPA] single," +
                           " [CHIPS Average Vol (ft3)] single," +
                           " [CHIPS Average Weight (tons)] single, " +
                           " [CHIPS Average Density (lbs/ft3)] single, " +
                           " [CHIPS Hwd Proportion] single," +
                           " [CHIPS Chip Fraction] single," +
                           " [CHIPS utilized logs (ft3)] single," +
                           " [CHIPS utilized chips (tons)] single," +
                           " [SMLOGS TPA] single," +
                           " [SMLOGS Average Vol (ft3)] single," +
                           " [SMLOGS Average Weight (tons)] single," +
                           " [SMLOGS Average Density (lbs/ft3)] single," +
                           " [SMLOGS Hwd Proportion] single," +
                           " [SMLOGS Chip Fraction] single," +
                           " [SMLOGS utilized logs (ft3)] single," +
                           " [SMLOGS utilized chips (tons)] single," +
                           " [LGLOGS TPA] single," +
                           " [LGLOGS Average Vol (ft3)] single," +
                           " [LGLOGS Average Weight (tons)] single," +
                           " [LGLOGS Average Density (lbs/ft3)] single," +
                           " [LGLOGS Hwd Proportion] single," +
                           " [LGLOGS Chip Fraction] single," +
                           " [LGLOGS utilized logs (ft3)] single," +
                           " [LGLOGS utilized chips (tons)] single," +
                           " [TOTAL TPA] single," +
                           " [TOTAL Average Vol (ft3)] single," +
                           " [TOTAL Average Weight (tons)] single," +
                           " [TOTAL Average Density (lbs/ft3)] single," +
                           " [TOTAL Hwd Proportion] single," +
                           " [TOTAL Chip Fraction] single," +
                           " [TOTAL utilized logs (ft3)] single," +
                           " [TOTAL utilized chips (tons)] single," +
                           " [BRUSH CUT utilized logs (ft3)] single," +
                           " [BRUSH CUT utilized chips (tons)] single," +
                           " [BRUSH CUT Residue at Stump Vol (ft3)] single," +
                           " [BRUSH CUT Residue at Landing Vol (ft3)] single," + 
                           " [BRUSH CUT not utilized TPA] single," + 
                           " [BRUSH CUT not utilized Average Vol (ft3)] single," + 
                           " [BRUSH CUT not utilized Average Weight (tons)] single, " +
                           " [BRUSH CUT not utilized Average Density (lbs/ft3)] single, " +
                           " [BRUSH CUT not utilized Hwd Proportion] single," +
                           " [BRUSH CUT not utilized Chip Fraction] single," +
                           " [BRUSH CUT not utilized logs (ft3)] single," +
                           " [BRUSH CUT not utilized chips (tons)] single)";
                
            }
            static public string CreateFRCSOutputTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                       "STAND CHAR(255)," +
                       "[$/Green Ton] CHAR(255)," +
                       "[$/CCF] CHAR(255)," +
                       "[$/Acre] CHAR(255)," +
                       "RxPackage_Rx_RxCycle CHAR(255))";
            }
            public void CreateFRCSWarningTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateFRCSWarningTableSQL(p_strTableName));
                CreateFRCSWarningTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateFRCSWarningTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "biosum_cond_id,rxpackage,rx,rxcycle");

            }
            static public string CreateFRCSWarningTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " " +
                          "(biosum_cond_id text (25)," +
                            "RXPackage text (3)," + 
                            "RX text (3)," +
                            "RXCycle text (1)," +
                            "lglogpa text (30)," +
                            "lglogpatoallpa text (30)," +
                            "chipvol text (30)," +
                            "lglogvol text (30)," +
                            "smlogvol text (30)," +
                            "alllogvol text (30)," +
                            "allvol text (30)," +
                            "slope text(30)," +
                            "warning text (220))";

            }


            public void CreateTotalAdditionalHarvestCostsTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateTotalAdditionalHarvestCostsTableSQL(p_strTableName));
                CreateTotalAdditionalHarvestCostsTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateTotalAdditionalHarvestCostsTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "biosum_cond_id,rx");

            }
            static public string CreateTotalAdditionalHarvestCostsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " " +
                          "(biosum_cond_id text (25)," +
                            "RX text (3)," +
                            "complete_additional_cpa double)";

            }



			
		}

		public class ProcessorScenarioRuleDefinitions
		{

			static public string DefaultTreeSpeciesDollarValuesDbFile {get {return @"db\scenario_processor_rule_definitions.mdb";}}
			static public string DefaultTreeSpeciesDollarValuesTableName {get {return "scenario_tree_species_diam_dollar_values";}}
			static public string DefaultRxHarvestMethodDbFile {get {return @"db\scenario_processor_rule_definitions.mdb";}}
			static public string DefaultRxHarvestMethodTableName {get {return "scenario_rx_harvest_method";}}
			static public string DefaultHarvestMethodDbFile {get {return @"db\scenario_processor_rule_definitions.mdb";}}
			static public string DefaultHarvestMethodTableName {get {return "scenario_harvest_method";}}
            static public string DefaultMoveInCostsDbFile { get { return @"db\scenario_processor_rule_definitions.mdb"; } }
            static public string DefaultMoveInCostsTableName { get { return "scenario_move_in_costs"; } }
            static public string DefaultHarvestCostColumnsDbFile { get { return @"db\scenario_processor_rule_definitions.mdb"; } }
            static public string DefaultHarvestCostColumnsTableName { get { return "scenario_harvest_cost_columns"; } }
			static public string DefaultCostRevenueEscalatorsDbFile {get {return @"db\scenario_processor_rule_definitions.mdb";}}
			static public string DefaultCostRevenueEscalatorsTableName {get {return "scenario_cost_revenue_escalators";}}
            static public string DefaultAdditionalHarvestCostsDbFile { get { return @"db\scenario_processor_rule_definitions.mdb"; } }
            static public string DefaultAdditionalHarvestCostsTableName { get { return "scenario_additional_harvest_costs"; } }
            static public string DefaultTreeDiamGroupsDbFile { get { return @"\db\scenario_processor_rule_definitions.mdb"; } }
            static public string DefaultTreeDiamGroupsTableName { get { return "scenario_tree_diam_groups"; } }
            static public string DefaultTreeSpeciesGroupsDbFile { get { return @"\db\scenario_processor_rule_definitions.mdb"; } }
            static public string DefaultTreeSpeciesGroupsTableName { get { return "scenario_tree_species_groups"; } }
            static public string DefaultTreeSpeciesGroupsListDbFile { get { return @"\db\scenario_processor_rule_definitions.mdb"; } }
            static public string DefaultTreeSpeciesGroupsListTableName { get { return "scenario_tree_species_groups_list"; } }


			
			
			public ProcessorScenarioRuleDefinitions()
			{
			}

			public void CreateScenarioTreeSpeciesDollarValuesTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateScenarioTreeSpeciesDollarValuesTableSQL(p_strTableName));
				CreateScenarioTreeSpeciesDollarValuesTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateScenarioTreeSpeciesDollarValuesTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				//p_oAdo.AddAutoNumber(p_oConn,p_strTableName,"id");
				//p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","scenario_id,id");
			}
			static public string CreateScenarioTreeSpeciesDollarValuesTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," + 
					"species_group INTEGER," + 
					"diam_group INTEGER," + 
                    "wood_bin CHAR(1) DEFAULT 'M'," + 
					"merch_value DECIMAL (4,2) DEFAULT 0," + 
					"chip_value DECIMAL (4,2) DEFAULT 0)";
			}
			public void CreateScenarioRxHarvestMethodTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,Tables.ProcessorScenarioRuleDefinitions.CreateScenarioRxHarvestMethodTableSQL(p_strTableName));
				CreateScenarioRxHarvestMethodTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateScenarioRxHarvestMethodTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","biosum_cond_id,rxpackage,rx,rxcycle");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_ScenarioId","scenario_id");
			}
			static public string CreateScenarioRxHarvestMethodTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," + 
					"biosum_cond_id CHAR(25)," + 
					"rxpackage CHAR(3)," + 
					"rx CHAR(3)," + 
					"rxcycle CHAR(1)," +
					"HarvestMethodId BYTE)";

			}
			public void CreateScenarioHarvestMethodTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,Tables.ProcessorScenarioRuleDefinitions.CreateScenarioHarvestMethodTableSQL(p_strTableName));
				CreateScenarioHarvestMethodTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateScenarioHarvestMethodTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_ScenarioId","scenario_id");
			}
			static public string CreateScenarioHarvestMethodTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," + 
					"HarvestMethodLowSlope CHAR(50)," +
					"HarvestMethodSteepSlope CHAR(50)," + 
                    "MaxCableYardingDistance SINGLE," + 
                    "MaxHelicopterCableYardingDistance SINGLE," + 
					"min_chip_dbh SINGLE," + 
					"min_sm_log_dbh SINGLE," + 
					"min_lg_log_dbh SINGLE," + 
					"SteepSlope INTEGER," + 
					"min_dbh_steep_slope SINGLE," + 
                    "ProcessLowSlopeYN CHAR(1) DEFAULT 'Y'," + 
                    "ProcessSteepSlopeYN CHAR(1) DEFAULT 'Y', " +
                    "WoodlandMerchAsPercentOfTotalVol INTEGER, " +
                    "SaplingMerchAsPercentOfTotalVol INTEGER, " +
                    "CullPctThreshold INTEGER, " +
                    "HarvestMethodSelection CHAR(15))";
			}

            public void CreateScenarioMoveInCostsTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, Tables.ProcessorScenarioRuleDefinitions.CreateScenarioMoveInCostsTableSQL(p_strTableName));
                CreateScenarioHarvestMethodTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateScenarioMoveInCostsTableIndexes(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_ScenarioId", "scenario_id");
            }
            static public string CreateScenarioMoveInCostsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "yard_dist_threshold SINGLE," +
                    "assumed_harvest_area_ac SINGLE," +
                    "move_in_time_multiplier SINGLE," +
                    "move_in_hours_addend SINGLE)"; 
            }

			public void CreateScenarioCostRevenueEscalatorsTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,Tables.ProcessorScenarioRuleDefinitions.CreateScenarioCostRevenueEscalatorsTableSQL(p_strTableName));
				CreateScenarioCostRevenueEscalatorsTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateScenarioCostRevenueEscalatorsTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_ScenarioId","scenario_id");
			}
			static public string CreateScenarioCostRevenueEscalatorsTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"scenario_id CHAR(20)," + 
					"EscalatorOperatingCosts_Cycle2 DECIMAL (4,2) DEFAULT 1.00," + 
					"EscalatorOperatingCosts_Cycle3 DECIMAL (4,2) DEFAULT 1.00," + 
					"EscalatorOperatingCosts_Cycle4 DECIMAL (4,2) DEFAULT 1.00," + 
					"EscalatorMerchWoodRevenue_Cycle2 DECIMAL (4,2) DEFAULT 1.00," + 
					"EscalatorMerchWoodRevenue_Cycle3 DECIMAL (4,2) DEFAULT 1.00," + 
					"EscalatorMerchWoodRevenue_Cycle4 DECIMAL (4,2) DEFAULT 1.00," + 
					"EscalatorEnergyWoodRevenue_Cycle2 DECIMAL (4,2) DEFAULT 1.00," + 
					"EscalatorEnergyWoodRevenue_Cycle3 DECIMAL (4,2) DEFAULT 1.00," + 
					"EscalatorEnergyWoodRevenue_Cycle4 DECIMAL (4,2) DEFAULT 1.00)";

			}
            public void CreateScenarioHarvestCostColumnsTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, Tables.ProcessorScenarioRuleDefinitions.CreateScenarioHarvestCostColumnsTableSQL(p_strTableName));
                CreateScenarioHarvestCostColumnsTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateScenarioHarvestCostColumnsTableIndexes(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_ScenarioId", "scenario_id");
            }
            static public string CreateScenarioHarvestCostColumnsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                     "scenario_id CHAR(20)," +
                     "ColumnName CHAR(50)," +
                     "rx CHAR(3)," + 
                     "Default_CPA DECIMAL (6,2)," + 
                     "Description CHAR(255))";

            }

            public void CreateScenarioAdditionalHarvestCostsTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, Tables.ProcessorScenarioRuleDefinitions.CreateScenarioAdditionalHarvestCostsTableSQL(p_strTableName));
                CreateScenarioAdditionalHarvestCostsTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateScenarioAdditionalHarvestCostsTableIndexes(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "scenario_id,biosum_cond_id,rx");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_plotrx", "biosum_cond_id,rx");
            }
            static public string CreateScenarioAdditionalHarvestCostsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "biosum_cond_id CHAR(25)," +
                    "rx CHAR(3))";
                    //"water_barring_roads_cpa DOUBLE," +
                    //"brush_cutting_cpa DOUBLE)";

            }

            public void CreateScenarioTreeDiamGroupsTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateScenarioTreeDiamGroupsTableSQL(p_strTableName));
                CreateScenarioTreeDiamGroupsTableIndexes(p_oAdo, p_oConn, p_strTableName);

            }
            public void CreateScenarioTreeDiamGroupsTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "diam_group,scenario_id");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_ScenarioId", "scenario_id");

            }
            public string CreateScenarioTreeDiamGroupsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "diam_group INTEGER DEFAULT 0," +
                    "diam_class CHAR(15)," +
                    "min_diam DOUBLE DEFAULT 0," +
                    "max_diam DOUBLE DEFAULT 0)";
            }
            public void CreateScenarioTreeSpeciesGroupsTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateScenarioTreeSpeciesGroupsTableSQL(p_strTableName));
                CreateScenarioTreeSpeciesGroupsTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateScenarioTreeSpeciesGroupsTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "species_group,scenario_id");
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_ScenarioId", "scenario_id");

            }
            public string CreateScenarioTreeSpeciesGroupsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "species_group INTEGER," +
                    "species_label CHAR(50)," +
                    "scenario_id CHAR(20))";
            }

            public void CreateScenarioTreeSpeciesGroupsListTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateScenarioTreeSpeciesGroupsListTableSQL(p_strTableName));
                CreateScenarioTreeSpeciesGroupsListTableIndexes(p_oAdo, p_oConn, p_strTableName);
            }
            public void CreateScenarioTreeSpeciesGroupsListTableIndexes(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "species_group");
            }
            public string CreateScenarioTreeSpeciesGroupsListTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "species_group INTEGER," +
                    "common_name CHAR(50), " +
                    "spcd INTEGER," +
                    "scenario_id CHAR(20))";
            }

		}
		public class Reference
		{
			public Reference()
			{
			}
			static public string DefaultTreeSpeciesTableDbFile {get {return @"db\ref_master.mdb";}}
			static public string DefaultTreeSpeciesTableName {get {return "tree_species";}}
			static public string DefaultOwnerGroupsTableDbFile {get {return @"db\ref_master.mdb";}}
			static public string DefaultOwnerGroupsTableName {get {return "owner_groups";}}
			static public string DefaultInventoriesTableDbFile {get {return @"db\ref_master.mdb";}}
			static public string DefaultInventoriesTableName {get {return "inventories";}}
			static public string DefaultFVSTreeSpeciesTableDbFile {get {return @"db\ref_master.mdb";}}
			static public string DefaultFVSTreeSpeciesTableName {get {return "fvs_tree_species";}}
			static public string DefaultFiadbFVSVariantTableDbFile {get {return @"db\ref_master.mdb";}}
			static public string DefaultFiadbFVSVariantTableName {get {return "fiadb_fvs_variant";}}
			static public string DefaultHarvestMethodsTableDbFile {get {return @"db\ref_master.mdb";}}
            static public string DefaultHarvestMethodsTableName { get { return "harvest_methods"; } }
            static public string DefaultRxCategoryTableDbFile { get { return @"db\ref_master.mdb"; } }
            static public string DefaultRxCategoryTableName { get { return "fvs_rx_category"; } }
            static public string DefaultRxSubCategoryTableDbFile { get { return @"db\ref_master.mdb"; } }
            static public string DefaultRxSubCategoryTableName { get { return "fvs_rx_subcategory"; } }
            static public string DefaultFVSCommandsTableDbFile { get { return @"db\ref_master.mdb"; } }
            static public string DefaultFVSCommandsTableName { get { return "fvs_commands"; } }
            static public string DefaultFVSWesternTreeSpeciesTableDbFile { get { return @"db\ref_master.mdb"; } }
            static public string DefaultFVSWesternTreeSpeciesTableName { get { return "FVS_WesternTreeSpeciesTranslator"; } }
            static public string DefaultFVSEasternTreeSpeciesTableDbFile { get { return @"db\ref_master.mdb"; } }
            static public string DefaultFVSEasternTreeSpeciesTableName { get { return "FVS_EasternTreeSpeciesTranslator"; } }
            static public string DefaultTreeMacroPlotBreakPointDiaTableDbFile { get { return @"db\ref_master.mdb"; } }
            static public string DefaultTreeMacroPlotBreakPointDiaTableName { get { return "TreeMacroPlotBreakPointDia"; } }
            static public string DefaultBiosumReferenceDbFile { get { return "biosum_ref.accdb"; } }
            static public string DefaultOpCostReferenceDbFile { get { return @"db\opcost_ref.accdb"; } }



			public void CreateTreeSpeciesTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateTreeSpeciesTableSQL(p_strTableName));
				CreateTreeSpeciesTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateTreeSpeciesTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","id");
				p_oAdo.AddAutoNumber(p_oConn,p_strTableName,"id");
			}
            static public string CreateTreeSpeciesTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                     "id LONG," +
                     "spcd INTEGER," +
                     "common_name CHAR(50)," +
                     "genus CHAR(20)," +
                     "species CHAR(50)," +
                     "variety CHAR(50)," +
                     "subspecies CHAR(50)," +
                     "fvs_variant CHAR(2)," +
                     "fvs_input_spcd INTEGER," +
                     "comments CHAR(200))";
            }
            public void CreateFVSCommandsTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateFVSCommandsTableSQL(p_strTableName));
				CreateFVSCommandsTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}

			public void CreateFVSCommandsTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				//p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","id");
				//p_oAdo.AddAutoNumber(p_oConn,p_strTableName,"id");
			}
			static public string CreateFVSCommandsTableSQL(string p_strTableName)
			{
			
			
				    return "CREATE TABLE " + p_strTableName + " (" +
                        "fvscmd_id INTEGER," +
                        "fvscmd CHAR(30)," + 
					    "[desc] MEMO," +
                        "fvs_variant_list CHAR(50)," +
                        "p1_label CHAR(50)," +
                        "p1_desc CHAR(100)," + 
					    "p1_default CHAR(10)," + 
					    "p1_validvalues CHAR(255)," + 
                        "p2_label CHAR(50)," +
                        "p2_desc CHAR(150)," + 
					    "p2_default CHAR(10)," + 
					    "p2_validvalues CHAR(255)," + 
                        "p3_label CHAR(50)," +
                        "p3_desc CHAR(200)," + 
					    "p3_default CHAR(10)," + 
					    "p3_validvalues CHAR(255)," + 
                        "p4_label CHAR(50)," +
                        "p4_desc CHAR(150)," + 
					    "p4_default CHAR(10)," + 
					    "p4_validvalues CHAR(255)," + 
                        "p5_label CHAR(50)," +
                        "p5_desc CHAR(150)," + 
					    "p5_default CHAR(10)," + 
					    "p5_validvalues CHAR(255)," + 
                        "p6_label CHAR(50)," +
                        "p6_desc CHAR(150)," + 
					    "p6_default CHAR(10)," + 
					    "p6_validvalues CHAR(255)," + 
                        "p7_label CHAR(50)," +
                        "p7_desc CHAR(255)," + 
					    "p7_default CHAR(10)," + 
					    "p7_validvalues CHAR(255)," + 
                        "other_label CHAR(50)," + 
                        "[other_desc] CHAR(255)," + 
                        "[other] MEMO," + 
                        "[filter] CHAR(100))";

			}
			public void CreateOwnerGroupsTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateOwnerGroupsTableSQL(p_strTableName));
				CreateOwnerGroupsTableIndexes(p_oAdo,p_oConn,p_strTableName);
				
				}
			public void CreateOwnerGroupsTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","owngrpcd");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","idb_owngrpcd");
			}
			static public string CreateOwnerGroupsTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"owngrpcd INTEGER," + 
					"idb_owngrpcd INTEGER," + 
					"`desc` CHAR(25))";
			}
			public void CreateInventoriesTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateInventoriesTableSQL(p_strTableName));
				CreateInventoriesTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateInventoriesTableIndexes(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx1","inv_id");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx2","idb_data_source");
				p_oAdo.AddIndex(p_oConn,p_strTableName,p_strTableName + "_idx3","inv_id_def");
			}
			static public string CreateInventoriesTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"inv_id CHAR(4)," + 
					"idb_data_source CHAR(6)," + 
					"description CHAR(50)," + 
					"inv_id_def CHAR(50)";
			}
			public void CreateFVSTreeSpeciesTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateFVSTreeSpeciesTableSQL(p_strTableName));
				CreateFVSTreeSpeciesTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateFVSTreeSpeciesTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				
			}
			static public string CreateFVSTreeSpeciesTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"spcd INTEGER," + 
					"common_name CHAR(50)," + 
					"species CHAR(50)," + 
					"genus CHAR(20)," + 
					"variety CHAR(50)," + 
					"subspecies CHAR(50)," + 
					"fvs_variant CHAR(2)," + 
					"fvs_species CHAR(2)," + 
					"fvs_common_name CHAR(50))";
			}
			public void CreateFiadbFVSVariantTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateFiadbFVSVariantTableSQL(p_strTableName));
				CreateFiadbFVSVariantTableIndexes(p_oAdo,p_oConn,p_strTableName);
			}
			public void CreateFiadbFVSVariantTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","statecd,countycd,plot");
			}
			static public string CreateFiadbFVSVariantTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"statecd INTEGER," + 
					"countycd INTEGER," +
					"plot LONG," + 
					"fvs_variant CHAR(2))";
			}

			public void CreateHarvestMethodsTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.SqlNonQuery(p_oConn,CreateHarvestMethodsTableSQL(p_strTableName));
			}
			public void CreateHarvestMethodsTableIndexes(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName)
			{
				p_oAdo.AddPrimaryKey(p_oConn,p_strTableName,p_strTableName + "_pk","HarvestMethodId");
			}
			
			static public string CreateHarvestMethodsTableSQL(string p_strTableName)
			{
				return "CREATE TABLE " + p_strTableName + " (" +
					"HarvestMethodId BYTE," + 
					"STEEP_YN CHAR(1)," + 
					"Method CHAR(50)," +
					"Description MEMO," +
                    "min_yard_distance_ft DOUBLE," +
                    "min_tpa DOUBLE," +
                    "min_avg_tree_vol_cf DOUBLE," +
                    "biosum_category INTEGER)";
			}

            public void CreateRxCategoryTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateRxCategoryTableSQL(p_strTableName));
            }
            public void CreateRxCategoryTableIndexes(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "catid");
            }

            static public string CreateRxCategoryTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "catid INTEGER," +
                    "[desc] CHAR(100)," +
                    "[min] INTEGER," +
                    "[max] INTEGER)";
            }
            public void CreateRxSubCategoryTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateRxSubCategoryTableSQL(p_strTableName));
            }
            public void CreateRxSubCategoryTableIndexes(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.AddPrimaryKey(p_oConn, p_strTableName, p_strTableName + "_pk", "catid");
            }

            static public string CreateRxSubCategoryTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "subcatid INTEGER," + 
                    "catid INTEGER," +
                    "[desc] CHAR(100)," +
                    "[min] INTEGER," +
                    "[max] INTEGER)";
            }
            public void CreateFVSWesternSpeciesTranslatorTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateFVSWesternSpeciesTranslatorTableSQL(p_strTableName));
            }
            static public string CreateFVSWesternSpeciesTranslatorTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "USDA_PLANTS_SYMBOL CHAR(10)," +
                    "FIA_SPCD CHAR(3)," +
                    "FVS_ALPHACODE CHAR(2)," +
                    "COMMON_NAME CHAR(50)," +
                    "SCIENTIFIC_NAME CHAR(50)," +
                    "AK_Mapped_To CHAR(2)," +
                    "BM_Mapped_To CHAR(2)," +
                    "CA_Mapped_To CHAR(2)," +
                    "CI_Mapped_To CHAR(2)," +
                    "CR_Mapped_To CHAR(2)," +
                    "EC_Mapped_To CHAR(2)," +
                    "EM_Mapped_To CHAR(2)," +
                    "IE_Mapped_To CHAR(2)," +
                    "KT_Mapped_To CHAR(2)," +
                    "NC_Mapped_To CHAR(2)," +
                    "NI_Mapped_To CHAR(2)," +
                    "PN_Mapped_To CHAR(2)," +
                    "SO_Mapped_To CHAR(2)," +
                    "TT_Mapped_To CHAR(2)," +
                    "UT_Mapped_To CHAR(2)," +
                    "WC_Mapped_To CHAR(2)," +
                    "WS_Mapped_To CHAR(2))"; 

            }
            public void CreateFVSEasternSpeciesTranslatorTable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
            {
                p_oAdo.SqlNonQuery(p_oConn, CreateFVSEasternSpeciesTranslatorTableSQL(p_strTableName));
            }
            static public string CreateFVSEasternSpeciesTranslatorTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "USDA_PLANTS_SYMBOL CHAR(10)," +
                    "FIA_SPCD CHAR(3)," +
                    "FVS_ALPHACODE CHAR(2)," +
                    "COMMON_NAME CHAR(50)," +
                    "SCIENTIFIC_NAME CHAR(50)," +
                    "CS_Mapped_To CHAR(2)," +
                    "LS_Mapped_To CHAR(2)," +
                    "NE_Mapped_To CHAR(2)," +
                    "SE_Mapped_To CHAR(2)," +
                    "SN_Mapped_To CHAR(2))";

            }
		}
	}
}
