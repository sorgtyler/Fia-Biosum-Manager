using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FIA_Biosum_Manager
{
    /// <summary>
    /// Objects and logic for processing cutlist into BioSum output
    /// </summary>
    public class processor
    {
        Queries m_oQueries = new Queries();
        RxTools m_oRxTools = new RxTools();
        //@ToDo: this will come from the UI
        private string _strScenarioId = "scenario1";
        private string _strOpcostTableName = "OPCOST_INPUT_NEW";
        private string m_strDebugFile = "";
        private ado_data_access m_oAdo;
        private List<tree> m_trees;
        private diameterVariables m_diamVariables;

        public processor(string strDebugFile)
        {
            m_strDebugFile = frmMain.g_oEnv.strTempDir + "\\" + strDebugFile;
        }
        
        public void init()
        {
            //
            //CREATE LINK IN TEMP MDB TO ALL PROCESSOR SCENARIO TABLES
            //
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "START: Create Links to the Scenario tables - " + System.DateTime.Now.ToString() + "\r\n");
            dao_data_access oDao = new dao_data_access();
            //
            //SCENARIO MDB
            //
            
            string strScenarioMDB =
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\db\\scenario_processor_rule_definitions.mdb";
            //
            //SCENARIO RESULTS MDB
            //
            string strScenarioResultsMDB =
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + _strScenarioId + "\\" + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsDbFile;

            //
            //LOAD PROJECT DATATASOURCES INFO
            //
            m_oQueries.m_oFvs.LoadDatasource = true;
            m_oQueries.m_oReference.LoadDatasource = true;
            m_oQueries.m_oProcessor.LoadDatasource = true;
            m_oQueries.LoadDatasources(true, "processor", _strScenarioId);

            //link to all the scenario rule definition tables
            oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
                "scenario_cost_revenue_escalators",
                strScenarioMDB, "scenario_cost_revenue_escalators", true);
            oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
                "scenario_additional_harvest_costs",
                strScenarioMDB, "scenario_additional_harvest_costs", true);
            oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
               "scenario_harvest_cost_columns",
               strScenarioMDB, "scenario_harvest_cost_columns", true);
            oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
              "scenario_harvest_method",
              strScenarioMDB, "scenario_harvest_method", true);
            oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
             "scenario_tree_species_diam_dollar_values",
             strScenarioMDB, "scenario_tree_species_diam_dollar_values", true);
            //link scenario results tables
            oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
                Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName,
                strScenarioResultsMDB,
                Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, true);
            oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
                Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName,
                strScenarioResultsMDB,
                Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName, true);

            oDao.m_DaoDbEngine.Idle(1);
            oDao.m_DaoDbEngine.Idle(8);
            oDao.m_DaoWorkspace.Close();
            oDao.m_DaoDbEngine = null;
            oDao = null;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "END: Create Links to the Scenario tables - " + System.DateTime.Now.ToString() + "\r\n");

            //
            //CREATE LINK IN TEMP MDB TO ALL VARIANT CUTLIST TABLES
            //
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "START: CreateTableLinksToFVSOutTreeListTables - " + System.DateTime.Now.ToString() + "\r\n");
            m_oRxTools.CreateTableLinksToFVSOutTreeListTables(m_oQueries, m_oQueries.m_strTempDbFile);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "END: CreateTableLinksToFVSOutTreeListTables - " + System.DateTime.Now.ToString() + "\r\n");

        }
        
        private void loadTrees(string p_strVariant, string p_strRxPackage)
        {
            //Load diameter groups into reference list
            List<treeDiamGroup> listDiamGroups = loadTreeDiamGroups();
            //Load species groups into reference dictionary
            IDictionary<string, int> dictSpeciesGroups = loadSpeciesGroups(p_strVariant);
            //Load species diam values into reference dictionary
            IDictionary<string, speciesDiamValue> dictSpeciesDiamValues = loadSpeciesDiamValues(_strScenarioId);
            //Load diameter variables into reference object
            m_diamVariables = loadDiameterVariables(_strScenarioId);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "loadTrees: Diameter Variables in Use: " + m_diamVariables.ToString() + "\r\n");


            
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));
            string strTableName = "fvs_tree_IN_" + p_strVariant + "_P" + p_strRxPackage + "_TREE_CUTLIST";
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT z.biosum_cond_id, z.rxCycle, z.rx, z.rxYear, " +
                                "z.dbh, z.tpa, z.volCfNet, z.drybiot, z.drybiom,z.FvsCreatedTree_YN, " +
                                "z.fvs_tree_id, z.fvs_species, " +
                                "c.slope, p.elev, p.gis_yard_dist " +
                                  "FROM " + strTableName + " z, cond c, plot p " +
                                  "WHERE z.rxpackage='" + p_strRxPackage + "' AND " +
                                  "z.biosum_cond_id = c.biosum_cond_id AND " +
                                  "c.biosum_plot_id = p.biosum_plot_id AND " +
                                  "mid(z.fvs_tree_id,1,2)='" + p_strVariant + "'";
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    m_trees = new List<tree>();
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        tree newTree = new tree();
                        newTree.CondId = Convert.ToString(m_oAdo.m_OleDbDataReader["biosum_cond_id"]).Trim();
                        newTree.RxCycle = Convert.ToString(m_oAdo.m_OleDbDataReader["rxCycle"]).Trim();
                        newTree.RxPackage = p_strRxPackage;
                        newTree.Rx= Convert.ToString(m_oAdo.m_OleDbDataReader["rx"]).Trim();
                        newTree.RxYear = Convert.ToString(m_oAdo.m_OleDbDataReader["rxYear"]).Trim();
                        newTree.Dbh = Convert.ToDouble(m_oAdo.m_OleDbDataReader["dbh"]);
                        newTree.Tpa = Convert.ToDouble(m_oAdo.m_OleDbDataReader["tpa"]);
                        newTree.VolCfNet = Convert.ToDouble(m_oAdo.m_OleDbDataReader["volCfNet"]);
                        newTree.DryBiot = Convert.ToDouble(m_oAdo.m_OleDbDataReader["drybiot"]);
                        newTree.DryBiom = Convert.ToDouble(m_oAdo.m_OleDbDataReader["drybiom"]);
                        newTree.Slope = Convert.ToInt32(m_oAdo.m_OleDbDataReader["slope"]);
                        newTree.FvsTreeId = Convert.ToString(m_oAdo.m_OleDbDataReader["fvs_tree_id"]).Trim();
                        if (Convert.ToString(m_oAdo.m_OleDbDataReader["FvsCreatedTree_YN"]).Trim().ToUpper() == "Y")
                        {
                            newTree.FvsCreatedTree = true;                           
                            // only use fvs_species from cut list if it is an FVS created tree
                            newTree.SpCd = Convert.ToString(m_oAdo.m_OleDbDataReader["fvs_species"]).Trim();
                        }
                        newTree.Elevation = Convert.ToInt32(m_oAdo.m_OleDbDataReader["elev"]);
                        newTree.YardingDistance = Convert.ToDouble(m_oAdo.m_OleDbDataReader["gis_yard_dist"]);
                        m_trees.Add(newTree);
                    }
                }

                //Query TREE table to get original FIA species codes
                strSQL = "SELECT DISTINCT t.fvs_tree_id, t.spcd " +
                        "FROM tree t, " + strTableName + " z " +
                        "WHERE t.fvs_tree_id = z.fvs_tree_id " +
                        "AND z.rxpackage='" + p_strRxPackage + "' " +
                        "AND mid(t.fvs_tree_id,1,2)='" + p_strVariant + "' " +
                        "GROUP BY t.fvs_tree_id, t.spcd";

                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    Dictionary<String, String> dictSpCd = new Dictionary<string, string>();
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        string strTreeId = Convert.ToString(m_oAdo.m_OleDbDataReader["fvs_tree_id"]).Trim();
                        string strSpCd = Convert.ToString(m_oAdo.m_OleDbDataReader["spcd"]).Trim();
                        dictSpCd.Add(strTreeId, strSpCd);
                    }

                    // Second pass at processing tree properties based on information from the cut list
                    foreach (tree nextTree in m_trees)
                    {
                        if (!nextTree.FvsCreatedTree)
                        {
                            nextTree.SpCd = dictSpCd[nextTree.FvsTreeId];
                        }
                        // set species group from species group dictionary
                        nextTree.SpeciesGroup = dictSpeciesGroups[nextTree.SpCd];

                        // set diameter group from diameter group list
                        foreach (treeDiamGroup nextGroup in listDiamGroups)
                        {
                            if (nextTree.Dbh >= nextGroup.MinDiam &&
                                nextTree.Dbh <= nextGroup.MaxDiam)
                            {
                                nextTree.DiamGroup = nextGroup.DiamGroup;
                                break;
                            }
                        }

                        // set tree properties based on scenario_tree_species_diam_dollar_values
                        string strSpeciesDiamKey = nextTree.DiamGroup + "|" + nextTree.SpeciesGroup;
                        speciesDiamValue treeSpeciesDiam = null;
                        if (dictSpeciesDiamValues.TryGetValue(strSpeciesDiamKey, out treeSpeciesDiam))
                        {
                            nextTree.MerchValue = treeSpeciesDiam.MerchValue;
                            nextTree.ChipValue = treeSpeciesDiam.ChipValue;
                            switch (treeSpeciesDiam.WoodBin)
                            {
                                case "M":
                                    nextTree.IsNonCommercial = false;
                                    break;
                                case "C":
                                    nextTree.IsNonCommercial = true;
                                    break;
                            }
                        }
                        else
                        {
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "loadTrees: Missing species diam values for diamGroup|speciesGroup " + 
                                    strSpeciesDiamKey + " - " + System.DateTime.Now.ToString() + "\r\n");
                        }

                        //Assign OpCostTreeType
                        nextTree.TreeType = chooseOpCostTreeType(nextTree);
                        //Dump OpCostTreeType in .csv format for validation
                        //string strLogEntry = nextTree.Dbh + ", " + nextTree.Slope + ", " + nextTree.IsNonCommercial +
                        //    ", " + nextTree.SpeciesGroup + ", " + nextTree.TreeType.ToString();
                        //if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        //    frmMain.g_oUtils.WriteText(m_strDebugFile, "loadTrees: OpCost tree type, " +
                        //        strLogEntry + "\r\n");


                        //if (nextTree.DiamGroup < 1)
                        //{
                        //    System.Windows.MessageBox.Show("missing diam group");
                        //}
                    }
                }
                //System.Windows.MessageBox.Show(m_trees.Count + " trees");
            }
            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;
        }

        private void createOpcostInput()
        {
            if (m_trees.Count < 1)
            {
                System.Windows.MessageBox.Show("No cut trees have been loaded for this scenario, variant, package combination. \r\n The OpCost input file cannot be created",
                    "FIA Biosum", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return;
            }

            // create connection to database
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));
            
            // drop opcost input table if it exists
            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, _strOpcostTableName) == true)
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + _strOpcostTableName);

            if (m_oAdo.m_intError == 0)
            {

                // create opcost input table
                frmMain.g_oTables.m_oProcessor.CreateOpcostInputTable(m_oAdo, m_oAdo.m_OleDbConnection, _strOpcostTableName);

                IDictionary<string, opcostInput> dictOpcostInput = new Dictionary<string, opcostInput>();
                foreach (tree nextTree in m_trees)
                {
                    opcostInput nextInput = null;
                    string strStand = nextTree.CondId + nextTree.RxPackage + nextTree.Rx + nextTree.RxCycle;
                    bool blnFound = dictOpcostInput.TryGetValue(strStand, out nextInput);
                    if (!blnFound)
                    {
                        //check helicopter system
                        //delete plots where yarding distance greater than 1300 feet if cable system
                        //m_oAdo.m_strSQL = "DELETE [One-way Yarding Distance], [Harvesting System] " +
                        //                  "FROM [" + p_strTable + "] " +
                        //                  "WHERE ((([One-way Yarding Distance])>" + ScenarioHarvestMethodVariables.MaxCableYardingDistance.ToString().Trim() + ") AND " +
                        //                         "(TRIM([Harvesting System]) IN ('Cable Manual WT/Log'," +
                        //                                                        "'Cable Manual WT'," +
                        //                                                        "'Cable Manual Log'," +
                        //                                                        "'Cable CTL')))";
                        
                        nextInput = new opcostInput(nextTree.CondId, nextTree.Slope, nextTree.RxCycle, nextTree.RxPackage,
                                                    nextTree.Rx, nextTree.RxYear, nextTree.YardingDistance, nextTree.Elevation);
                        dictOpcostInput.Add(strStand, nextInput);
                    }
                }
                System.Windows.MessageBox.Show(dictOpcostInput.Keys.Count + " lines in file");
            }
            
            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;
        }

        private List<treeDiamGroup> loadTreeDiamGroups()
        {
            List<treeDiamGroup> listDiamGroups = new List<treeDiamGroup>();
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT * FROM tree_diam_groups";
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        int intDiamGroup = Convert.ToInt32(m_oAdo.m_OleDbDataReader["diam_group"]);
                        double dblMinDiam = Convert.ToDouble(m_oAdo.m_OleDbDataReader["min_diam"]);
                        double dblMaxDiam = Convert.ToDouble(m_oAdo.m_OleDbDataReader["max_diam"]);
                        listDiamGroups.Add(new treeDiamGroup(intDiamGroup, dblMinDiam, dblMaxDiam));
                    }
                }
            }
            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;

            return listDiamGroups;
        }

        private IDictionary<String, int> loadSpeciesGroups(string p_strVariant)
        {
            IDictionary<String, int> dictSpeciesGroups = new Dictionary<String, int>();
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT DISTINCT SPCD, USER_SPC_GROUP FROM tree_species " +
                                "WHERE FVS_VARIANT = '" + p_strVariant + "' " +
                                "AND SPCD IS NOT NULL " +
                                "AND USER_SPC_GROUP IS NOT NULL " +
                                "GROUP BY SPCD, USER_SPC_GROUP";
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        string strSpCd = Convert.ToString(m_oAdo.m_OleDbDataReader["SPCD"]).Trim();
                        int intSpcGroup = Convert.ToInt32(m_oAdo.m_OleDbDataReader["USER_SPC_GROUP"]);
                        dictSpeciesGroups.Add(strSpCd, intSpcGroup);
                    }
                }
            }
            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;

            return dictSpeciesGroups;
        }

        ///<summary>
        /// Loads scenario_tree_species_diam_dollar_values into a reference dictionary
        /// The composite key is intDiamGroup + "|" + intSpcGroup
        /// The value is a speciesDiamValue object
        ///</summary> 
        private IDictionary<String, speciesDiamValue> loadSpeciesDiamValues(string p_scenario)
        {
            IDictionary<String, speciesDiamValue> dictSpeciesDiamValues = new Dictionary<String, speciesDiamValue>();
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT * FROM scenario_tree_species_diam_dollar_values " +
                                "WHERE scenario_id = '" + p_scenario + "'";
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        int intSpcGroup = Convert.ToInt32(m_oAdo.m_OleDbDataReader["species_group"]);
                        int intDiamGroup = Convert.ToInt32(m_oAdo.m_OleDbDataReader["diam_group"]);
                        string strWoodBin = Convert.ToString(m_oAdo.m_OleDbDataReader["wood_bin"]).Trim();
                        double dblMerchValue = Convert.ToDouble(m_oAdo.m_OleDbDataReader["merch_value"]);
                        double dblChipValue = Convert.ToDouble(m_oAdo.m_OleDbDataReader["chip_value"]);
                        string strKey = intDiamGroup + "|" + intSpcGroup;
                        dictSpeciesDiamValues.Add(strKey, new speciesDiamValue(intDiamGroup, intSpcGroup,
                            strWoodBin, dblMerchValue, dblChipValue));
                    }
                    //Console.WriteLine("DiamValues: " + dictSpeciesDiamValues.Keys.Count);
                }
            }
            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;

            return dictSpeciesDiamValues;
        }

        private diameterVariables loadDiameterVariables(string p_scenario)
        {
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));
            diameterVariables returnVariables = null;
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT * FROM scenario_harvest_method " +
                                "WHERE scenario_id = '" + p_scenario + "'";
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    // We should only have one record
                    m_oAdo.m_OleDbDataReader.Read();
                    double dblMinChipDbh = Convert.ToDouble(m_oAdo.m_OleDbDataReader["min_chip_dbh"]);
                    double dblMinSmallLogDbh = Convert.ToDouble(m_oAdo.m_OleDbDataReader["min_sm_log_dbh"]);
                    double dblMinLgLogDbh = Convert.ToDouble(m_oAdo.m_OleDbDataReader["min_lg_log_dbh"]);
                    int intMinSlopePct = Convert.ToInt32(m_oAdo.m_OleDbDataReader["SteepSlope"]);
                    double dblMinDbhSteepSlope = Convert.ToDouble(m_oAdo.m_OleDbDataReader["min_dbh_steep_slope"]);

                    returnVariables = new diameterVariables(dblMinChipDbh, dblMinSmallLogDbh, dblMinLgLogDbh,
                        intMinSlopePct, dblMinDbhSteepSlope);

                    //Console.WriteLine("DiamValues: " + dictSpeciesDiamValues.Keys.Count);
                }
            }

            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;

            return returnVariables;
        }
        
        private OpCostTreeType chooseOpCostTreeType(tree p_tree)
        {
            OpCostTreeType returnType = OpCostTreeType.None;
            if (p_tree.Dbh < m_diamVariables.MinChipDbh)
            {
                returnType = OpCostTreeType.BC;
            }
            else if (p_tree.Slope >= m_diamVariables.SteepSlopePct && 
                     p_tree.Dbh < m_diamVariables.MinDbhSteepSlope)
            {
                returnType = OpCostTreeType.BC;
            }
            else if (p_tree.IsNonCommercial)
            {
                returnType = OpCostTreeType.CT;
            }
            else if (p_tree.Dbh >= m_diamVariables.MinChipDbh && 
                     p_tree.Dbh < m_diamVariables.MinSmallLogDbh)
            {
                returnType = OpCostTreeType.CT;
            }
            else if (p_tree.Dbh >= m_diamVariables.MinSmallLogDbh &&
                     p_tree.Dbh < m_diamVariables.MinLargeLogDbh)
            {
                returnType = OpCostTreeType.SL;
            }
            else if (p_tree.Dbh >= m_diamVariables.MinLargeLogDbh)
            {
                returnType = OpCostTreeType.LL;
            }

            return returnType;
        }

        enum OpCostTreeType
        {
            None = 0,
            BC = 1,
            CT = 2,
            SL = 3,
            LL = 4
        };
        
        ///<summary>
        ///Represents a tree in the fvs cutlist
        ///</summary>
        private class tree
        {
            string _strCondId = "";
            string _strRxCycle = "";
            string _strRxPackage = "";
            string _strRx = "";
            string _strRxYear = "";
            double _dblDbh;
            double _dblTpa;
            double _dblVolCfNet;
            double _dblDryBiot;
            double _dblDryBiom;
            int _intSlope;
            string _strSpcd;
            bool _boolFvsCreatedTree;
            string _strFvsTreeId;
            OpCostTreeType _TreeType;
            int _intSpeciesGroup;
            int _intDiamGroup;
            bool _boolIsNonCommercial;
            double _dblMerchValue;
            double _dblChipValue;
            int _intElev;
            double _dblYardingDistance;

            string _strDebugFile;

            public tree()
			{
               
			}

            public string CondId
            {
                get { return _strCondId; }
                set { _strCondId = value; }
            }
            public string RxCycle
            {
                get { return _strRxCycle; }
                set { _strRxCycle = value; }
            }
            public string RxPackage
            {
                get { return _strRxPackage; }
                set { _strRxPackage = value; }
            }            
            public string Rx
            {
                get { return _strRx; }
                set { _strRx = value; }
            }            
            public string RxYear
            {
                get { return _strRxYear; }
                set { _strRxYear = value; }
            }
            public double Dbh
            {
                get { return _dblDbh; }
                set { _dblDbh = value; }
            }
            public double Tpa
            {
                get { return _dblTpa; }
                set { _dblTpa = value; }
            }
            public double VolCfNet
            {
                get { return _dblVolCfNet; }
                set { _dblVolCfNet = value; }
            }
            public double DryBiot
            {
                get { return _dblDryBiot; }
                set { _dblDryBiot = value; }
            }
            public double DryBiom
            {
                get { return _dblDryBiom; }
                set { _dblDryBiom = value; }
            }
            public int Slope
            {
                get { return _intSlope; }
                set { _intSlope = value; }
            }
            public string SpCd
            {
                get { return _strSpcd; }
                set { _strSpcd = value; }
            }
            public bool FvsCreatedTree
            {
                get { return _boolFvsCreatedTree; }
                set { _boolFvsCreatedTree = value; }
            }
            public string FvsTreeId
            {
                get { return _strFvsTreeId; }
                set { _strFvsTreeId = value; }
            }
            public OpCostTreeType TreeType
            {
                get { return _TreeType; }
                set { _TreeType = value; }
            }
            public int SpeciesGroup
            {
                get { return _intSpeciesGroup; }
                set { _intSpeciesGroup = value; }
            }
            public int DiamGroup
            {
                get { return _intDiamGroup; }
                set { _intDiamGroup = value; }
            }
            public bool IsNonCommercial
            {
                get { return _boolIsNonCommercial; }
                set { _boolIsNonCommercial = value; }
            }
            public double MerchValue
            {
                get { return _dblMerchValue; }
                set { _dblMerchValue = value; }
            }
            public double ChipValue
            {
                get { return _dblChipValue; }
                set { _dblChipValue = value; }
            }
            public int Elevation
            {
                get { return _intElev; }
                set { _intElev = value; }
            }
            public double YardingDistance
            {
                get { return _dblYardingDistance; }
                set { _dblYardingDistance = value; }
            }
        }

        private class treeDiamGroup
        {
            int _intDiamGroup;
            double _dblMinDiam;
            double _dblMaxDiam;

            public treeDiamGroup(int diamGroup, double dblMinDiam, double dblMaxDiam)
			{
                _intDiamGroup = diamGroup;
                _dblMinDiam = dblMinDiam;
                _dblMaxDiam = dblMaxDiam;
			}

            public int DiamGroup
            {
                get { return _intDiamGroup; }
            }
            public double MinDiam
            {
                get { return _dblMinDiam; }
            }
            public double MaxDiam
            {
                get { return _dblMaxDiam; }
            }
        }

        private class speciesDiamValue
        {
            int _intSpeciesGroup;
            int _intDiamGroup;
            string _strWoodBin;
            double _dblMerchValue;
            double _dblChipValue;

            public speciesDiamValue(int diamGroup, int speciesGroup, string woodBin, double merchValue, double chipValue)
			{
                _intDiamGroup = diamGroup;
                _intSpeciesGroup = speciesGroup;
                _strWoodBin = woodBin;
                _dblMerchValue = merchValue;
                _dblChipValue = chipValue;
			}

            public int DiamGroup
            {
                get { return _intDiamGroup; }
            }
            public int SpeciesGroup
            {
                get { return _intSpeciesGroup; }
            }
            public string WoodBin
            {
                get { return _strWoodBin; }
            }
            public double MerchValue
            {
                get { return _dblMerchValue; }
            }
            public double ChipValue
            {
                get { return _dblChipValue; }
            }
        }

        private class diameterVariables
        {
            double _dblMinSmallLogDbh;
            double _dblMinLargeLogDbh;
            double _dblMinChipDbh;
            int _intSteepSlopePct;
            double _dblMinDbhSteepSlope;

            public diameterVariables(double minChipDbh, double minSmallLogDbh, double minLargeLogDbh, int steepSlopePct,
                                     double minDbhSteepSlope)
            {
                _dblMinSmallLogDbh = minSmallLogDbh;
                _dblMinLargeLogDbh = minLargeLogDbh;
                _dblMinChipDbh = minChipDbh;
                _intSteepSlopePct = steepSlopePct;
                _dblMinDbhSteepSlope = minDbhSteepSlope;
            }

            public double MinChipDbh
            {
                get { return _dblMinChipDbh; }
            }
            public double MinSmallLogDbh
            {
                get { return _dblMinSmallLogDbh; }
            }
            public double MinLargeLogDbh
            {
                get { return _dblMinLargeLogDbh; }
            }
            public int SteepSlopePct
            {
                get { return _intSteepSlopePct; }
            }
            public double MinDbhSteepSlope
            {
                get { return _dblMinDbhSteepSlope; }
            }
            // Overriding the ToString method for debugging purposes
            public override string ToString()
            {
                return string.Format("MinChipDbh: {0}, MinSmallLogDbh: {1}, MinLargeLogDbh: {2}, SteepSlopePct: {3}, MinDbhSteepSlope: {4}]", 
                    _dblMinChipDbh, _dblMinSmallLogDbh, _dblMinLargeLogDbh, _intSteepSlopePct, _dblMinDbhSteepSlope);
            }
        }

        /// <summary>
        /// An opcostInput object represents a line in the opcostInput file
        /// The metrics are aggregated by stand with is a unique concatenation of
        /// conditionId, rxPackage, rx, and rxCycle
        /// </summary>
        private class opcostInput
        {
            string _strCondId = "";
            int _intPercentSlope;
            string _strRxCycle = "";
            string _strRxPackage = "";
            string _strRx = "";
            string _strRxYear = "";
            double _dblYardingDistance;
            int _intElev;

            public opcostInput(string condId, int percentSlope, string rxCycle, string rxPackage, string rx,
                               string rxYear, double yardingDistance, int elev)
            {
                _strCondId = condId;
                _intPercentSlope = percentSlope;
                _strRxCycle = rxCycle;
                _strRxPackage = rxPackage;
                _strRx = rx;
                _strRxYear = rxYear;
                _dblYardingDistance = yardingDistance;
                _intElev = elev;
            }

            public string OpCostStand    
            {
                get { return _strCondId + _strRxPackage + _strRx + _strRxCycle; }
            }
            public int PercentSlope
            {
                get { return _intPercentSlope; }
            }
            public double YardingDistance
            {
                get { return _dblYardingDistance; }
            }
            public string RxYear
            {
                get { return _strRxYear; }
            }
            public int Elev
            {
                get { return _intElev; }
            }
            public string RxPackageRxRxCycle
            {
                get { return _strRxPackage + _strRx + _strRxCycle; }
            }
        }

    }
}
