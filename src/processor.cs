using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
    /// <summary>
    /// Objects and logic for processing cutlist into BioSum output
    /// </summary>
    public class processor
    {
        private string m_strScenarioId = "";
        private string m_strTempDbFile = "";
        private string m_strOpcostTableName = "opcost_input";
        private string m_strTvvTableName = "TreeVolValLowSlope";
        private string m_strDebugFile =
            "";
        private ado_data_access m_oAdo;
        private List<tree> m_trees;
        private scenarioHarvestMethod m_scenarioHarvestMethod;
        private IDictionary<string, prescription> m_prescriptions;
        private IList<harvestMethod> m_harvestMethodList;
        private escalators m_escalators;

        public processor(string strDebugFile, string strScenarioId)
        {
            m_strDebugFile = strDebugFile;
            m_strScenarioId = strScenarioId;
        }
        
        public Queries init()
        {
            Queries p_oQueries = new Queries();
            
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
                "\\processor\\" + m_strScenarioId + "\\" + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsDbFile;

            //
            //LOAD PROJECT DATATASOURCES INFO
            //
            p_oQueries.m_oFvs.LoadDatasource = true;
            p_oQueries.m_oReference.LoadDatasource = true;
            p_oQueries.m_oProcessor.LoadDatasource = true;
            p_oQueries.LoadDatasources(true, "processor", m_strScenarioId);

            //link to all the scenario rule definition tables
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
                "scenario_cost_revenue_escalators",
                strScenarioMDB, "scenario_cost_revenue_escalators", true);
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
                "scenario_additional_harvest_costs",
                strScenarioMDB, "scenario_additional_harvest_costs", true);
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
               "scenario_harvest_cost_columns",
               strScenarioMDB, "scenario_harvest_cost_columns", true);
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
              "scenario_harvest_method",
              strScenarioMDB, "scenario_harvest_method", true);
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
             "scenario_tree_species_diam_dollar_values",
             strScenarioMDB, "scenario_tree_species_diam_dollar_values", true);
            //link scenario results tables
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
                Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName,
                strScenarioResultsMDB,
                Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, true);
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
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
            RxTools oRxTools = new RxTools();
            oRxTools.CreateTableLinksToFVSOutTreeListTables(p_oQueries, p_oQueries.m_strTempDbFile);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "END: CreateTableLinksToFVSOutTreeListTables - " + System.DateTime.Now.ToString() + "\r\n");

            return p_oQueries;
        }
        
        public void loadTrees(string p_strVariant, string p_strRxPackage, string strTempDbFile)
        {
            //Set tempDbFile
            m_strTempDbFile = strTempDbFile;
            //Load harvest methods; Prescription load depends on harvest methods
            m_harvestMethodList = loadHarvestMethods();
            //Load prescriptions into reference dictionary
            m_prescriptions = loadPrescriptions(strTempDbFile);
            //Load diameter variables into reference object
            m_scenarioHarvestMethod = loadScenarioHarvestMethod(m_strScenarioId);
            //Load escalators into reference object
            m_escalators = loadEscalators();

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "loadTrees: Diameter Variables in Use: " + m_scenarioHarvestMethod.ToString() + "\r\n");
            
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(strTempDbFile, "", ""));
            string strTableName = "fvs_tree_IN_" + p_strVariant + "_P" + p_strRxPackage + "_TREE_CUTLIST";
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT z.biosum_cond_id, z.rxCycle, z.rx, z.rxYear, " +
                                "z.dbh, z.tpa, z.volCfNet, z.drybiot, z.drybiom,z.FvsCreatedTree_YN, " +
                                "z.fvs_tree_id, z.fvs_species, z.volTsGrs, " +
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
                        newTree.DebugFile = m_strDebugFile;
                        newTree.CondId = Convert.ToString(m_oAdo.m_OleDbDataReader["biosum_cond_id"]).Trim();
                        newTree.RxCycle = Convert.ToString(m_oAdo.m_OleDbDataReader["rxCycle"]).Trim();
                        newTree.RxPackage = p_strRxPackage;
                        newTree.Rx= Convert.ToString(m_oAdo.m_OleDbDataReader["rx"]).Trim();
                        newTree.RxYear = Convert.ToString(m_oAdo.m_OleDbDataReader["rxYear"]).Trim();
                        newTree.Dbh = Convert.ToDouble(m_oAdo.m_OleDbDataReader["dbh"]);
                        newTree.Tpa = Convert.ToDouble(m_oAdo.m_OleDbDataReader["tpa"]);
                        newTree.VolCfNet = Convert.ToDouble(m_oAdo.m_OleDbDataReader["volCfNet"]);
                        newTree.VolTsGrs = Convert.ToDouble(m_oAdo.m_OleDbDataReader["volTsGrs"]);
                        newTree.DryBiot = Convert.ToDouble(m_oAdo.m_OleDbDataReader["drybiot"]);
                        newTree.DryBiom = Convert.ToDouble(m_oAdo.m_OleDbDataReader["drybiom"]);
                        newTree.Slope = Convert.ToInt32(m_oAdo.m_OleDbDataReader["slope"]);
                        // find default harvest methods in prescription in case we need them
                        string strDefaultHarvestMethodLowSlope = "";
                        string strDefaultHarvestMethodSteepSlope = "";
                        int intDefaultHarvestMethodCategoryLowSlope = 0;
                        int intDefaultHarvestMethodCategorySteepSlope = 0 ;
                        prescription currentPrescription = null;
                        m_prescriptions.TryGetValue(newTree.Rx, out currentPrescription);
                        if (currentPrescription != null)
                        {
                            strDefaultHarvestMethodLowSlope = currentPrescription.HarvestMethodLowSlope;
                            strDefaultHarvestMethodSteepSlope = currentPrescription.HarvestMethodSteepSlope;
                            intDefaultHarvestMethodCategoryLowSlope = currentPrescription.HarvestCategoryLowSlope;
                            intDefaultHarvestMethodCategorySteepSlope = currentPrescription.HarvestCategorySteepSlope;
                        }

                        if (newTree.Slope < m_scenarioHarvestMethod.SteepSlopePct)
                        {
                            // assign low slope harvest method
                            if (! String.IsNullOrEmpty(m_scenarioHarvestMethod.HarvestMethodLowSlope))
                            {
                                newTree.HarvestMethod = m_scenarioHarvestMethod.HarvestMethodLowSlope;
                                newTree.HarvestMethodCategory = m_scenarioHarvestMethod.HarvestCategoryLowSlope;
                            }
                            else
                            {
                                newTree.HarvestMethod = strDefaultHarvestMethodLowSlope;
                                newTree.HarvestMethodCategory = intDefaultHarvestMethodCategoryLowSlope;
                            }
                        }
                        else
                        {
                            // assign steep slope harvest method
                            if (!String.IsNullOrEmpty(m_scenarioHarvestMethod.HarvestMethodSteepSlope))
                            {
                                newTree.HarvestMethod = m_scenarioHarvestMethod.HarvestMethodSteepSlope;
                                newTree.HarvestMethodCategory = m_scenarioHarvestMethod.HarvestCategorySteepSlope;
                            }
                            else
                            {
                                newTree.HarvestMethod = strDefaultHarvestMethodSteepSlope;
                                newTree.HarvestMethodCategory = intDefaultHarvestMethodCategorySteepSlope;
                            }
                        }
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
            }
            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;
        }

        public void updateTrees(string p_strVariant, string p_strRxPackage, string strTempDbFile)
        {
            if (m_trees.Count < 1)
            {
                System.Windows.MessageBox.Show("No cut trees have been loaded for this scenario, variant, package combination. \r\n Auxillary tree data cannot be appended",
                    "FIA Biosum", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return;
            }

            m_strTempDbFile = strTempDbFile;

            //Load species groups into reference dictionary
            IDictionary<string, treeSpecies> dictTreeSpecies = loadTreeSpecies(p_strVariant);
            //Load species diam values into reference dictionary
            IDictionary<string, speciesDiamValue> dictSpeciesDiamValues = loadSpeciesDiamValues(m_strScenarioId);
            //Load diameter groups into reference list
            List<treeDiamGroup> listDiamGroups = loadTreeDiamGroups();
            
            //Query TREE table to get original FIA species codes
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(strTempDbFile, "", ""));

            if (m_oAdo.m_intError == 0)
            {
            string strTableName = "fvs_tree_IN_" + p_strVariant + "_P" + p_strRxPackage + "_TREE_CUTLIST";
            string strSQL = "SELECT DISTINCT t.fvs_tree_id, t.spcd " +
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
                    // set tree species fields from treeSpecies dictionary
                    treeSpecies foundSpecies = dictTreeSpecies[nextTree.SpCd];
                    nextTree.SpeciesGroup = foundSpecies.SpeciesGroup;
                    nextTree.OdWgt = foundSpecies.OdWgt;
                    nextTree.DryToGreen = foundSpecies.DryToGreen;

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
                    calculateVolumeAndWeight(nextTree);

                    //Dump OpCostTreeType in .csv format for validation
                    //string strLogEntry = nextTree.Dbh + ", " + nextTree.Slope + ", " + nextTree.IsNonCommercial +
                    //    ", " + nextTree.SpeciesGroup + ", " + nextTree.TreeType.ToString();
                    //if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    //    frmMain.g_oUtils.WriteText(m_strDebugFile, "loadTrees: OpCost tree type, " +
                    //        strLogEntry + "\r\n");
                }
            }
            }
            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;
        }
        
        public void createOpcostInput(string strTempDbFile)
        {
            m_strTempDbFile = strTempDbFile;
            int intHwdSpeciesCodeThreshold = 299; // Species codes greater than this are hardwoods
            if (m_trees.Count < 1)
            {
                System.Windows.MessageBox.Show("No cut trees have been loaded for this scenario, variant, package combination. \r\n The OpCost input file cannot be created",
                    "FIA Biosum", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return;
            }

            // create connection to database
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(strTempDbFile, "", ""));
            
            // drop opcost input table if it exists
            //if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, m_strOpcostTableName) == true)
            //    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + m_strOpcostTableName);

            if (m_oAdo.m_intError == 0)
            {

                // create opcost input table
                frmMain.g_oTables.m_oProcessor.CreateOpcostInputTable(m_oAdo, m_oAdo.m_OleDbConnection, m_strOpcostTableName);


                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createOpcostInput: Read trees into opcost input - " + System.DateTime.Now.ToString() + "\r\n");

                IDictionary<string, opcostInput> dictOpcostInput = new Dictionary<string, opcostInput>();
                foreach (tree nextTree in m_trees)
                {
                    // if the tree yarding distance exceeds the user maximum, don't process it
                    if (!nextTree.exceedsYardingLimit(m_scenarioHarvestMethod.MaxCableYardingDistance, m_scenarioHarvestMethod.MaxHelicopterCableYardingDistance))
                    {
                        opcostInput nextInput = null;
                        string strStand = nextTree.CondId + nextTree.RxPackage + nextTree.Rx + nextTree.RxCycle;
                        bool blnFound = dictOpcostInput.TryGetValue(strStand, out nextInput);
                        if (!blnFound)
                        {
                            nextInput = new opcostInput(nextTree.CondId, nextTree.Slope, nextTree.RxCycle, nextTree.RxPackage,
                                                    nextTree.Rx, nextTree.RxYear, nextTree.YardingDistance, nextTree.Elevation,
                                                    nextTree.HarvestMethod);
                            dictOpcostInput.Add(strStand, nextInput);
                        }
                    
                        // Metrics for brush cut trees
                        if (nextTree.TreeType == OpCostTreeType.BC)
                        {
                            nextInput.TotalBcTpa = nextInput.TotalBcTpa + nextTree.Tpa;
                            //nextInput.TotalBcVolCf = nextInput.TotalBcVolCf + nextTree.OpCostBrushCutVolCf;
                        }
                        // Metrics for chip trees
                        else if (nextTree.TreeType == OpCostTreeType.CT)
                        {
                            nextInput.TotalChipTpa = nextInput.TotalChipTpa + nextTree.Tpa;
                            //nextInput.TotalCtTreeBiomass = nextInput.TotalCtTreeBiomass + nextTree.DryBiot;
                            nextInput.TotalChipMerchVolCf = nextInput.TotalChipMerchVolCf + nextTree.DryBiom;
                            //nextInput.TotalCtVolCf = nextInput.TotalCtVolCf + nextTree.OpCostChipVolCf;
                            //nextInput.TotalCtWtGt = nextInput.TotalCtWtGt + nextTree.OpCostChipWtGt;
                            //if (Convert.ToInt32(nextTree.SpCd) > intHwdSpeciesCodeThreshold)
                                //nextInput.TotalCtHwdVolCf = nextInput.TotalCtHwdVolCf + nextTree.OpCostChipVolCf;

                        }
                        // Metrics for small log trees
                        else if (nextTree.TreeType == OpCostTreeType.SL)
                        {
                            nextInput.TotalSmLogTpa = nextInput.TotalSmLogTpa + nextTree.Tpa;
                            //nextInput.TotalSmLogTreeBiomass = nextInput.TotalSmLogTreeBiomass + nextTree.DryBiot;
                            //nextInput.TotalSmLogBoleBiomass = nextInput.TotalSmLogBoleBiomass + nextTree.DryBiom;
                            //nextInput.TotalSmLogVolCf = nextInput.TotalSmLogVolCf + nextTree.OpCostMerchVolCf;
                            //nextInput.TotalSmLogWtGt = nextInput.TotalSmLogWtGt + nextTree.OpCostMerchWtGt;
                            //if (Convert.ToInt32(nextTree.SpCd) > intHwdSpeciesCodeThreshold)
                            //    nextInput.TotalSmLogHwdVolCf = nextInput.TotalSmLogHwdVolCf + nextTree.OpCostMerchVolCf;
                        }
                        // Metrics for small log trees
                        else if (nextTree.TreeType == OpCostTreeType.LL)
                        {
                            nextInput.TotalLgLogTpa = nextInput.TotalLgLogTpa + nextTree.Tpa;
                            //nextInput.TotalLgLogTreeBiomass = nextInput.TotalLgLogTreeBiomass + nextTree.DryBiot;
                            //nextInput.TotalLgLogBoleBiomass = nextInput.TotalLgLogBoleBiomass + nextTree.DryBiom;
                            //nextInput.TotalLgLogVolCf = nextInput.TotalLgLogVolCf + nextTree.OpCostMerchVolCf;
                            //nextInput.TotalLgLogWtGt = nextInput.TotalLgLogWtGt + nextTree.OpCostMerchWtGt;
                            //if (Convert.ToInt32(nextTree.SpCd) > intHwdSpeciesCodeThreshold)
                            //    nextInput.TotalLgLogHwdVolCf = nextInput.TotalLgLogHwdVolCf + nextTree.OpCostMerchVolCf;
                        }
                    }
                }
                //System.Windows.MessageBox.Show(dictOpcostInput.Keys.Count + " lines in file");

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createOpcostInput: Finished reading trees - " + System.DateTime.Now.ToString() + "\r\n");

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createOpcostInput: Begin writing opcost input table - " + System.DateTime.Now.ToString() + "\r\n");
                long lngCount =0;
                foreach (string key in dictOpcostInput.Keys)
                {                    
                    opcostInput nextStand = dictOpcostInput[key];

                    // Some fields we wait to calculate until we have the totals
                    // *** BRUSH CUT ***
                    double dblBcAvgVolume = 0;

                    if (nextStand.TotalBcTpa > 0)
                        { dblBcAvgVolume = nextStand.TotalBcVolCf / nextStand.TotalBcTpa; }

                    // *** CHIP TREES ***
                    double dblCtResidueFraction = 0;
                    //if (nextStand.TotalChipMerchVolCf > 0)
                    //    { dblCtResidueFraction = Math.Round((nextStand.TotalCtTreeBiomass - nextStand.TotalChipMerchVolCf) / nextStand.TotalChipMerchVolCf) * 100; }
                    double dblCtAvgVolume = 0;
                    if (nextStand.TotalChipTpa > 0)
                        { dblCtAvgVolume = nextStand.TotalChipVolCf / nextStand.TotalChipTpa; }
                    double dblCtAvgDensity = 0;
                    double dblCtHwdProp = 0;
                    if (nextStand.TotalChipVolCf > 0)
                    {
                        dblCtAvgDensity = nextStand.TotalChipWtGt * 2000 / nextStand.TotalChipVolCf;
                        dblCtHwdProp = nextStand.TotalChipHwdVolCf / nextStand.TotalChipVolCf;
                    }

                    
                    // *** SMALL LOGS ***
                    double dblSmLogResidueFraction = 0;
                    //if (nextStand.TotalSmLogBoleBiomass > 0)
                    //    { dblSmLogResidueFraction = Math.Round((nextStand.TotalSmLogTreeBiomass - nextStand.TotalSmLogBoleBiomass) / nextStand.TotalSmLogBoleBiomass) * 100; }
                    double dblSmLogAvgVolume = 0;
                    if (nextStand.TotalSmLogTpa > 0)
                        { dblSmLogAvgVolume = nextStand.TotalSmLogVolCf / nextStand.TotalSmLogTpa; }
                    double dblSmLogAvgDensity = 0;
                    double dblSmLogHwdProp = 0;
                    if (nextStand.TotalSmLogVolCf > 0)
                    { 
                        dblSmLogAvgDensity = nextStand.TotalSmLogWtGt * 2000 / nextStand.TotalSmLogVolCf;
                        dblSmLogHwdProp = nextStand.TotalSmLogHwdVolCf / nextStand.TotalSmLogVolCf; 
                    }

                    // *** LARGE LOGS ***
                    double dblLgLogResidueFraction = 0;
                    //if (nextStand.TotalLgLogBoleBiomass > 0)
                    //{ dblLgLogResidueFraction = Math.Round((nextStand.TotalLgLogTreeBiomass - nextStand.TotalLgLogBoleBiomass) / nextStand.TotalLgLogBoleBiomass) * 100; }
                    double dblLgLogAvgVolume = 0;
                    if (nextStand.TotalLgLogTpa > 0)
                    { dblLgLogAvgVolume = nextStand.TotalLgLogVolCf / nextStand.TotalLgLogTpa; }
                    double dblLgLogAvgDensity = 0;
                    double dblLgLogHwdProp = 0;
                    if (nextStand.TotalLgLogVolCf > 0)
                    {
                        dblLgLogAvgDensity = nextStand.TotalLgLogWtGt * 2000 / nextStand.TotalLgLogVolCf;
                        dblLgLogHwdProp = nextStand.TotalLgLogHwdVolCf / nextStand.TotalLgLogVolCf;
                    }
  
                    m_oAdo.m_strSQL = "INSERT INTO " + m_strOpcostTableName + " " +
                    "(Stand, [Percent Slope], [One-way Yarding Distance], YearCostCalc, " +
                    "[Project Elevation], [Harvesting System], [Chip tree per acre], [Residue fraction for chip trees], " +
                    "[Chip trees average volume(ft3)], [CHIPS Average Density (lbs/ft3)], [CHIPS Hwd Proportion], [Small log trees per acre],  " +
                    "[Small log trees residue fraction], [Small log trees average volume(ft3)], [Small log trees average density(lbs/ft3)], " +
                    "[Small log trees hardwood proportion], [Large log trees per acre], [Large log trees residue fraction], " +
                    "[Large log trees average vol(ft3)], [Large log trees average density(lbs/ft3)], [Large log trees hardwood proportion], " +
                    "BrushCutTPA, [BrushCutAvgVol], RxPackage_Rx_RxCycle) " +
                    "VALUES ('" + nextStand.OpCostStand + "', " + nextStand.PercentSlope + ", " + nextStand.YardingDistance + ", '" + nextStand.RxYear + "', " +
                    nextStand.Elev + ", '" + nextStand.HarvestSystem + "', " + nextStand.TotalChipTpa + ", " + 
                    dblCtResidueFraction + ", " + dblCtAvgVolume + ", " + dblCtAvgDensity + ", " + dblCtHwdProp + ", " +
                    nextStand.TotalSmLogTpa + ", " + dblSmLogResidueFraction + ", " +
                    dblSmLogAvgVolume + ", " + dblSmLogAvgDensity + ", " + dblSmLogHwdProp + ", " +
                    nextStand.TotalLgLogTpa + ", " + dblLgLogResidueFraction + ", " + dblLgLogAvgVolume + ", " +
                    dblLgLogAvgDensity + ", " + dblLgLogHwdProp + "," + nextStand.TotalBcTpa + ", " + dblBcAvgVolume + 
                    ",'" + nextStand.RxPackageRxRxCycle + "' )";

                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    if (m_oAdo.m_intError != 0) break;
                    lngCount++;

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "END createOpcostInput INSERTED " + lngCount + " RECORDS: " + System.DateTime.Now.ToString() + "\r\n");

                }
            }
            
            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;
        }

        public void createTreeVolValWorkTable(string strDateTimeCreated, string strTempDbFile, bool blnInclHarvMethodCat)
        {
            if (m_trees.Count < 1)
            {
                System.Windows.MessageBox.Show("No cut trees have been loaded for this scenario, variant, package combination. \r\n The tree vol val cannot be created",
                    "FIA Biosum", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return;
            }

            
            // create connection to database
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(strTempDbFile, "", ""));

            if (m_oAdo.m_intError == 0)
            {
                // create tree vol val work table (TreeVolValLowSlope); Re-use the sql from tree vol val but don't create the indexes
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, Tables.Processor.CreateTreeVolValSpeciesDiamGroupsTableSQL(m_strTvvTableName));

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createTreeVolValWorkTable: Read trees into tree vol val - " + System.DateTime.Now.ToString() + "\r\n");

                string strSeparator = "_";
                IDictionary<string, treeVolValInput> dictTvvInput = new Dictionary<string, treeVolValInput>();
                foreach (tree nextTree in m_trees)
                {
                    treeVolValInput nextInput = null;
                    string strKey = nextTree.CondId + strSeparator + nextTree.RxCycle + strSeparator + nextTree.DiamGroup + strSeparator + nextTree.SpeciesGroup;
                    bool blnFound = dictTvvInput.TryGetValue(strKey, out nextInput);
                    if (!blnFound)
                    {
                        //calculate chipMktValPgt; Apply escalators according to rxCycle
                        double chipMktValPgt = nextTree.ChipValue;
                        switch (nextTree.RxCycle)
                        {
                            case "2":
                                chipMktValPgt = chipMktValPgt * m_escalators.EnergyWoodRevCycle2;
                                break;
                            case "3":
                                chipMktValPgt = chipMktValPgt * m_escalators.EnergyWoodRevCycle3;
                                break;
                            case "4":
                                chipMktValPgt = chipMktValPgt * m_escalators.EnergyWoodRevCycle4;
                                break;
                        }
                        nextInput = new treeVolValInput(nextTree.CondId, nextTree.RxCycle, nextTree.RxPackage, nextTree.Rx,
                            nextTree.SpeciesGroup, nextTree.DiamGroup, nextTree.IsNonCommercial, chipMktValPgt, nextTree.HarvestMethodCategory);
                        dictTvvInput.Add(strKey, nextInput);
                    }

                    // Metrics for brush cut trees
                    if (nextTree.TreeType == OpCostTreeType.BC)
                    {
                        nextInput.TotalBrushCutWtGt = nextInput.TotalBrushCutWtGt + nextTree.BrushCutWtGt;
                        nextInput.TotalBrushCutVolumeCf = nextInput.TotalBrushCutVolumeCf + nextTree.BrushCutVolCf;
                        //nextInput.TotalStandResidueWtGt = nextInput.TotalStandResidueWtGt + nextTree.StandResidueWtGt;
                    }

                }
                
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createTreeVolValWorkTable: Finished reading trees - " + System.DateTime.Now.ToString() + "\r\n");

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createTreeVolValWorkTable: Begin writing tree vol val table - " + System.DateTime.Now.ToString() + "\r\n");
                long lngCount =0;
                foreach (string key in dictTvvInput.Keys)
                {
                    treeVolValInput nextStand = dictTvvInput[key];
                    m_oAdo.m_strSQL = "INSERT INTO " + m_strTvvTableName + " " +
                    "(biosum_cond_id, rxpackage, rx, rxcycle, species_group, diam_group, " +
                    "merch_to_chipbin_YN, chip_mkt_val_pgt, bc_vol_cf, bc_wt_gt, stand_residue_wt_gt, " +
                    "biosum_harvest_method_category, DateTimeCreated, place_holder)" +
                    "VALUES ('" + nextStand.CondId + "', '" + nextStand.RxPackage + "', '" + nextStand.Rx + "', '" + 
                    nextStand.RxCycle + "', " + nextStand.SpeciesGroup + ", " + nextStand.DiamGroup +
                    ", '" + nextStand.MerchToChip + "', " + nextStand.ChipMktValPgt + ", " + nextStand.TotalBrushCutVolumeCf + "," +
                    nextStand.TotalBrushCutWtGt + ", " + nextStand.TotalStandResidueWtGt + ", " + 
                    nextStand.HarvestMethodCategory + ", '" + strDateTimeCreated + "', 'N')";

                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    if (m_oAdo.m_intError != 0) break;
                    lngCount++;
                }

                //We may want this column for testing but not in the final product
                //Also drop id column because it prevents copying rows into final tree vol val
                if (!blnInclHarvMethodCat)
                {
                    string strSqlAlter = "ALTER TABLE " + m_strTvvTableName + " DROP COLUMN biosum_harvest_method_category, id";
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, strSqlAlter);
                }
                
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "END createTreeVolValWorkTable INSERTED " + lngCount + " RECORDS: " + System.DateTime.Now.ToString() + "\r\n");
            }
        }

        private List<treeDiamGroup> loadTreeDiamGroups()
        {
            List<treeDiamGroup> listDiamGroups = new List<treeDiamGroup>();
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strTempDbFile, "", ""));
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT * FROM " + Tables.Processor.DefaultTreeDiamGroupsTableName;
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

        private IDictionary<String, treeSpecies> loadTreeSpecies(string p_strVariant)
        {
            IDictionary<String, treeSpecies> dictTreeSpecies = new Dictionary<String, treeSpecies>();
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strTempDbFile, "", ""));
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT DISTINCT SPCD, USER_SPC_GROUP, OD_WGT, Dry_to_Green FROM " + 
                                Tables.Reference.DefaultTreeSpeciesTableName +
                                " WHERE FVS_VARIANT = '" + p_strVariant + "' " +
                                "AND SPCD IS NOT NULL " +
                                "AND USER_SPC_GROUP IS NOT NULL " +
                                "GROUP BY SPCD, USER_SPC_GROUP, OD_WGT, Dry_to_Green";
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        string strSpCd = Convert.ToString(m_oAdo.m_OleDbDataReader["SPCD"]).Trim();
                        int intSpcGroup = Convert.ToInt32(m_oAdo.m_OleDbDataReader["USER_SPC_GROUP"]);
                        double dblOdWgt = Convert.ToDouble(m_oAdo.m_OleDbDataReader["OD_WGT"]);
                        double dblDryToGreen = Convert.ToDouble(m_oAdo.m_OleDbDataReader["Dry_to_Green"]);
                        treeSpecies nextTreeSpecies = new treeSpecies(strSpCd, intSpcGroup, dblOdWgt, dblDryToGreen);
                        dictTreeSpecies.Add(strSpCd, nextTreeSpecies);
                    }
                }
            }
            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;

            return dictTreeSpecies;
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
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strTempDbFile, "", ""));
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT * FROM " + 
                                Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName + 
                                " WHERE scenario_id = '" + p_scenario + "'";
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

        private IDictionary<String, prescription> loadPrescriptions(string strTempDbFile)
        {
            if (m_harvestMethodList == null || m_harvestMethodList.Count == 0)
            {
                MessageBox.Show("Harvest methods must be loaded before loading prescriptions", "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
            IDictionary<String, prescription> dictPrescriptions = new Dictionary<String, prescription>();
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(strTempDbFile, "", ""));
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT * FROM " + Tables.FVS.DefaultRxTableName;
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        string strRx = Convert.ToString(m_oAdo.m_OleDbDataReader["rx"]).Trim();
                        string strHarvestMethodLowSlope = Convert.ToString(m_oAdo.m_OleDbDataReader["HarvestMethodLowSlope"]).Trim();
                        string strHarvestMethodSteepSlope = Convert.ToString(m_oAdo.m_OleDbDataReader["HarvestMethodSteepSlope"]).Trim();
                        int intHarvestCategoryLowSlope = 0;
                        int intHarvestCategorySteepSlope = 0;
                        foreach (harvestMethod nextMethod in m_harvestMethodList)
                        {
                            if (nextMethod.Method.Equals(strHarvestMethodLowSlope) && !nextMethod.SteepSlope)
                            {
                                intHarvestCategoryLowSlope = nextMethod.BiosumCategory;
                            }
                            else if (nextMethod.Method.Equals(strHarvestMethodSteepSlope) && nextMethod.SteepSlope)
                            {
                                intHarvestCategorySteepSlope = nextMethod.BiosumCategory;
                            }
                        }
                        
                        dictPrescriptions.Add(strRx, new prescription(strRx, strHarvestMethodLowSlope, strHarvestMethodSteepSlope, 
                            intHarvestCategoryLowSlope, intHarvestCategorySteepSlope));
                    }
                }
            }
            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;

            return dictPrescriptions;
        }

        private scenarioHarvestMethod loadScenarioHarvestMethod(string p_scenario)
        {
            if (m_harvestMethodList == null || m_harvestMethodList.Count == 0)
            {
                MessageBox.Show("Harvest methods must be loaded before loading scenario harvest methods", "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strTempDbFile, "", ""));
            scenarioHarvestMethod returnVariables = null;
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName +
                                " WHERE scenario_id = '" + p_scenario + "'";
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
                    double dblMaxCableYardingDistance = Convert.ToDouble(m_oAdo.m_OleDbDataReader["MaxCableYardingDistance"]);
                    double dblMaxHelicopterCableYardingDistance = Convert.ToDouble(m_oAdo.m_OleDbDataReader["MaxHelicopterCableYardingDistance"]);
                    string strHarvestMethodLowSlope = Convert.ToString(m_oAdo.m_OleDbDataReader["HarvestMethodLowSlope"]).Trim();
                    string strHarvestMethodSteepSlope = Convert.ToString(m_oAdo.m_OleDbDataReader["HarvestMethodSteepSlope"]).Trim();
                    int intGenericMerchAsPercentOfTotalVol = Convert.ToInt16(m_oAdo.m_OleDbDataReader["GenericMerchAsPercentOfTotalVol"]);
                    int intHarvestCategoryLowSlope = 0;
                    int intHarvestCategorySteepSlope = 0;
                    foreach (harvestMethod nextMethod in m_harvestMethodList)
                    {
                        if (nextMethod.Method.Equals(strHarvestMethodLowSlope) && !nextMethod.SteepSlope)
                        {
                            intHarvestCategoryLowSlope = nextMethod.BiosumCategory;
                        }
                        else if (nextMethod.Method.Equals(strHarvestMethodSteepSlope) && nextMethod.SteepSlope)
                        {
                            intHarvestCategorySteepSlope = nextMethod.BiosumCategory;
                        }
                    }
                    
                    returnVariables = new scenarioHarvestMethod(dblMinChipDbh, dblMinSmallLogDbh, dblMinLgLogDbh,
                        intMinSlopePct, dblMinDbhSteepSlope, dblMaxCableYardingDistance, dblMaxHelicopterCableYardingDistance,
                        strHarvestMethodLowSlope, strHarvestMethodSteepSlope, intHarvestCategoryLowSlope, intHarvestCategorySteepSlope,
                        intGenericMerchAsPercentOfTotalVol);
                }
            }

            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;

            return returnVariables;
        }

        private escalators loadEscalators()
        {
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strTempDbFile, "", ""));
            escalators returnEscalators = null;
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT * FROM " +
                                Tables.ProcessorScenarioRuleDefinitions.DefaultCostRevenueEscalatorsTableName + 
                                " WHERE scenario_id = '" + m_strScenarioId + "'";
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    // We should only have one record
                    m_oAdo.m_OleDbDataReader.Read();
                    double dblEnergyWoodRevCycle2 = Convert.ToDouble(m_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle2"]);
                    double dblEnergyWoodRevCycle3 = Convert.ToDouble(m_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle3"]);
                    double dblEnergyWoodRevCycle4 = Convert.ToDouble(m_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle4"]);
                    double dblMerchWoodRevCycle2 = Convert.ToDouble(m_oAdo.m_OleDbDataReader["EscalatorMerchWoodRevenue_Cycle2"]);
                    double dblMerchWoodRevCycle3 = Convert.ToDouble(m_oAdo.m_OleDbDataReader["EscalatorMerchWoodRevenue_Cycle3"]);
                    double dblMerchWoodRevCycle4 = Convert.ToDouble(m_oAdo.m_OleDbDataReader["EscalatorMerchWoodRevenue_Cycle4"]);


                    returnEscalators = new escalators(dblEnergyWoodRevCycle2, dblEnergyWoodRevCycle3, dblEnergyWoodRevCycle4,
                                                      dblMerchWoodRevCycle2, dblMerchWoodRevCycle3, dblMerchWoodRevCycle4);
                }
            }

            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;

            return returnEscalators;
        }

        private IList<harvestMethod> loadHarvestMethods()
        {
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strTempDbFile, "", ""));
            IList<harvestMethod> harvestMethodList = null;

            if (m_oAdo.m_intError == 0)
            {
                // Check to see if the biosum_category column exists in the harvest method table; If not
                // throw an error and exit the function; Processor won't work without this value
                if (!m_oAdo.ColumnExist(m_oAdo.m_OleDbConnection, Tables.Reference.DefaultFRCSHarvestMethodTableName, "biosum_category"))
                {
                    string strErrMsg = "Your project contains an obsolete version of the " + Tables.Reference.DefaultFRCSHarvestMethodTableName +
                                       " table that does not contain the 'biosum_category' field. Copy a new version of this table into your project from the" +
                                       " BioSum installation directory before trying to run Processor.";
                    MessageBox.Show(strErrMsg,"FIA Biosum",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    return harvestMethodList;
                }

                string strSQL = "SELECT * FROM " + Tables.Reference.DefaultFRCSHarvestMethodTableName;
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    harvestMethodList = new List<harvestMethod>();
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        string strSteepYN = Convert.ToString(m_oAdo.m_OleDbDataReader["STEEP_YN"]).Trim();
                        bool blnSteep = false;
                        if (strSteepYN.Equals("Y"))
                            { blnSteep = true; }
                        string strMethod = Convert.ToString(m_oAdo.m_OleDbDataReader["Method"]).Trim();
                        int intBiosumCategory = Convert.ToInt16(m_oAdo.m_OleDbDataReader["biosum_category"]);
                        harvestMethod newMethod = new harvestMethod(blnSteep, strMethod, intBiosumCategory);
                        harvestMethodList.Add(newMethod);
                    }
                }
            }

            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;

            return harvestMethodList;
        }
        
        private OpCostTreeType chooseOpCostTreeType(tree p_tree)
        {
            OpCostTreeType returnType = OpCostTreeType.None;

            if (p_tree.Dbh < m_scenarioHarvestMethod.MinChipDbh)
            {
                returnType = OpCostTreeType.BC;
            }
            else if (p_tree.Slope >= m_scenarioHarvestMethod.SteepSlopePct && 
                     p_tree.Dbh < m_scenarioHarvestMethod.MinDbhSteepSlope)
            {
                returnType = OpCostTreeType.BC;
            }
            else if (p_tree.IsNonCommercial)
            {
                returnType = OpCostTreeType.CT;
            }
            else if (p_tree.Dbh >= m_scenarioHarvestMethod.MinChipDbh && 
                     p_tree.Dbh < m_scenarioHarvestMethod.MinSmallLogDbh)
            {
                returnType = OpCostTreeType.CT;
            }
            else if (p_tree.Dbh >= m_scenarioHarvestMethod.MinSmallLogDbh &&
                     p_tree.Dbh < m_scenarioHarvestMethod.MinLargeLogDbh)
            {
                returnType = OpCostTreeType.SL;
            }
            else if (p_tree.Dbh >= m_scenarioHarvestMethod.MinLargeLogDbh)
            {
                returnType = OpCostTreeType.LL;
            }

            return returnType;
        }

        private void calculateVolumeAndWeight(tree p_tree)
        {
            int intGenericMerchFactor = m_scenarioHarvestMethod.GenericMerchAsPercentOfTotalVol / 100;
            //adjDryBiom - Do this first; precursor to other calculations
            if (p_tree.DryBiot <= p_tree.DryBiom)
            {
                p_tree.AdjDryBiom = intGenericMerchFactor * p_tree.DryBiom;
                
            }
            
            //merchVolCf
            if (p_tree.VolCfNet > 0)
            {
                p_tree.MerchVolCf = p_tree.VolCfNet * p_tree.Tpa;
            }
            else
            {
                p_tree.MerchVolCf = p_tree.VolTsGrs * intGenericMerchFactor * p_tree.Tpa;
            }

            //merchValDpa
            if (!p_tree.IsNonCommercial)
            {
                switch (p_tree.RxCycle)
                {
                    case "1":
                        p_tree.MerchValDpa = p_tree.MerchVolCf * p_tree.MerchValue;
                        return;
                    case "2":
                        p_tree.MerchValDpa = p_tree.MerchVolCf * p_tree.MerchValue * m_escalators.MerchWoodRevCycle2;
                        return;
                    case "3":
                        p_tree.MerchValDpa = p_tree.MerchVolCf * p_tree.MerchValue * m_escalators.MerchWoodRevCycle3;
                        return;
                    case "4":
                        p_tree.MerchValDpa = p_tree.MerchVolCf * p_tree.MerchValue * m_escalators.MerchWoodRevCycle4;
                        return;
                }
            }

            //merchWtGt
            p_tree.MerchWtGt = p_tree.VolCfNet * p_tree.OdWgt * p_tree.Tpa / p_tree.DryToGreen / 2000;

            //nonMerchVolCf
            if (p_tree.AdjDryBiom > 0)
            {
                p_tree.NonMerchVolCf = ((p_tree.DryBiot - p_tree.AdjDryBiom) * p_tree.Tpa) / p_tree.OdWgt;
            }
            else
            {
                p_tree.NonMerchVolCf = ((p_tree.DryBiot - p_tree.DryBiom) * p_tree.Tpa) / p_tree.OdWgt;
            }

            //nonMerchWtGt
            if (p_tree.AdjDryBiom > 0)
            {
                p_tree.NonMerchWtGt = ((p_tree.DryBiot - p_tree.AdjDryBiom) * p_tree.Tpa) / p_tree.OdWgt;
            }
            else
            {
                p_tree.NonMerchVolCf = ((p_tree.DryBiot - p_tree.DryBiom) * p_tree.Tpa) / p_tree.OdWgt;
            }

            //brushcut
            p_tree.BrushCutVolCf = p_tree.DryBiot * p_tree.Tpa / p_tree.OdWgt;
            p_tree.BrushCutWtGt = (p_tree.DryBiot * p_tree.Tpa / p_tree.DryToGreen) / 2000;
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
            double _dblVolTsGrs;
            double _dblDryBiot;
            double _dblDryBiom;
            int _intSlope;
            string _strSpcd;
            bool _boolFvsCreatedTree;
            string _strFvsTreeId;
            OpCostTreeType _opCostTreeType;
            int _intSpeciesGroup;
            int _intDiamGroup;
            bool _boolIsNonCommercial;
            double _dblMerchValue;
            double _dblChipValue;
            int _intElev;
            double _dblYardingDistance;
            string _strHarvestMethod;
            int _intHarvestMethodCategory;
            double _dblOdWgt;
            double _dblDryToGreen;
            double _dblMerchVolCf;
            double _dblMerchWtGt;
            double _dblBrushCutVolCf;
            double _dblBrushCutWtGt;
            double _dblNonMerchVolCf;
            double _dblNonMerchWtGt;
            double _dblAdjDryBiom;
            double _dblMerchValDpa;

            string _strDebugFile = "";

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
            public double VolTsGrs
            {
                get { return _dblVolTsGrs; }
                set { _dblVolTsGrs = value; }
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
                get { return _opCostTreeType; }
                set { _opCostTreeType = value; }
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
            public double OdWgt
            {
                get { return _dblOdWgt; }
                set { _dblOdWgt = value; }
            }
            public double DryToGreen
            {
                get { return _dblDryToGreen; }
                set { _dblDryToGreen = value; }
            }
            public double MerchVolCf
            {
                get { return _dblMerchVolCf; }
                set { _dblMerchVolCf = value; }
            }
            public double MerchWtGt
            {
                get { return _dblMerchWtGt; }
                set { _dblMerchWtGt = value; }
            }
            public double BrushCutVolCf
            {
                get { return _dblBrushCutVolCf; }
                set { _dblBrushCutVolCf = value; }
            }
            public double BrushCutWtGt
            {
                get { return _dblBrushCutWtGt; }
                set { _dblBrushCutWtGt = value; }
            }
            public double NonMerchVolCf
            {
                get { return _dblNonMerchVolCf; }
                set { _dblNonMerchVolCf = value; }
            }
            public double NonMerchWtGt
            {
                get { return _dblNonMerchWtGt; }
                set { _dblNonMerchWtGt = value; }
            }
            public string HarvestMethod
            {
                get { return _strHarvestMethod; }
                set { _strHarvestMethod = value; }
            }
            public int HarvestMethodCategory
            {
                get { return _intHarvestMethodCategory; }
                set { _intHarvestMethodCategory = value; }
            }
            public double AdjDryBiom
            {
                get { return _dblAdjDryBiom; }
                set { _dblAdjDryBiom = value; }
            }
            public double MerchValDpa
            {
                get { return _dblMerchValDpa; }
                set { _dblMerchValDpa = value; }
            }
            public double TotalVolCf
            {
                get 
                {
                    if (_opCostTreeType != OpCostTreeType.BC)
                    {
                        return _dblMerchVolCf + _dblNonMerchVolCf;
                    }
                    else
                    {
                        return _dblBrushCutVolCf;
                    }
                }
            }
            public double TotalVolCf
            {
                get
                {
                    if (_opCostTreeType != OpCostTreeType.BC)
                    {
                        return _dblMerchWtGt + _dblNonMerchWtGt;
                    }
                    else
                    {
                        return _dblBrushCutWtGt;
                    }
                }
            }
            public string DebugFile
            {
                set { _strDebugFile = value; }
            }

            public bool exceedsYardingLimit(double maxCableYardingDistance, double maxHelicopterYardingDistance)
            {
                if (string.IsNullOrEmpty(_strHarvestMethod))
                {
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(_strDebugFile, "tree.exceedsYardingLimit: No harvest method assigned \r\n");
                    return false;
                }
                if (_dblYardingDistance < 1)
                {
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(_strDebugFile, "tree.exceedsYardingLimit: Invalid yarding distance: " + _dblYardingDistance + " \r\n");
                    return false;
                }
                if (_strHarvestMethod.Contains("Cable") && _dblYardingDistance > maxCableYardingDistance)
                {
                    //Cable harvest method exceeds cable yarding limit
                    return true;
                }
                else if (_strHarvestMethod.Contains("Helicopter") && _dblYardingDistance > maxHelicopterYardingDistance)
                {
                    //Helicopter harvest method exceeds helicopter yarding limit
                    return true;
                }
                else {return false;}
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

        private class scenarioHarvestMethod
        {
            double _dblMinSmallLogDbh;
            double _dblMinLargeLogDbh;
            double _dblMinChipDbh;
            int _intSteepSlopePct;
            double _dblMinDbhSteepSlope;
            double _dblMaxCableYardingDistance;
            double _dblMaxHelicopterCableYardingDistance;
            string _strHarvestMethodLowSlope;
            string _strHarvestMethodSteepSlope;
            int _intHarvestCategoryLowSlope;
            int _intHarvestCategorySteepSlope;
            int _intGenericMerchAsPercentOfTotalVol;

            public scenarioHarvestMethod(double minChipDbh, double minSmallLogDbh, double minLargeLogDbh, int steepSlopePct,
                                         double minDbhSteepSlope, double maxCableYardingDistance, double maxHelicopterYardingDistance,
                                         string harvestMethodLowSlope, string harvestMethodSteepSlope,
                                         int harvestCategoryLowSlope, int harvestCategorySteepSlope, int genericMerchAsPercentOfTotalVol)
            {
                _dblMinSmallLogDbh = minSmallLogDbh;
                _dblMinLargeLogDbh = minLargeLogDbh;
                _dblMinChipDbh = minChipDbh;
                _intSteepSlopePct = steepSlopePct;
                _dblMinDbhSteepSlope = minDbhSteepSlope;
                _dblMaxCableYardingDistance = maxCableYardingDistance;
                _dblMaxHelicopterCableYardingDistance = maxHelicopterYardingDistance;
                _strHarvestMethodLowSlope = harvestMethodLowSlope;
                _strHarvestMethodSteepSlope = harvestMethodSteepSlope;
                _intHarvestCategoryLowSlope = harvestCategoryLowSlope;
                _intHarvestCategorySteepSlope = harvestCategorySteepSlope;
                _intGenericMerchAsPercentOfTotalVol = genericMerchAsPercentOfTotalVol;
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
            public double MaxCableYardingDistance
            {
                get { return _dblMaxCableYardingDistance; }
            }
            public double MaxHelicopterCableYardingDistance
            {
                get { return _dblMaxHelicopterCableYardingDistance; }
            }
            public string HarvestMethodLowSlope
            {
                get { return _strHarvestMethodLowSlope; }
            }
            public string HarvestMethodSteepSlope
            {
                get { return _strHarvestMethodSteepSlope; }
            }
            public int HarvestCategoryLowSlope
            {
                get { return _intHarvestCategoryLowSlope; }
            }
            public int HarvestCategorySteepSlope
            {
                get { return _intHarvestCategorySteepSlope; }
            }
            public int GenericMerchAsPercentOfTotalVol
            {
                get { return _intGenericMerchAsPercentOfTotalVol; }
            }
            // Overriding the ToString method for debugging purposes
            public override string ToString()
            {
                return string.Format("MinChipDbh: {0}, MinSmallLogDbh: {1}, MinLargeLogDbh: {2}, SteepSlopePct: {3}, MinDbhSteepSlope: {4}, " +
                    "MaxCableYardingDistance: {5}, MaxHelicopterCableYardingDistance: {6}, HarvestMethodLowSlope: {7}, HarvestMethodSteepSlope: {8}, " +
                    "GenericMerchAsPercentOfTotalVol: {9} ]",
                    _dblMinChipDbh, _dblMinSmallLogDbh, _dblMinLargeLogDbh, _intSteepSlopePct, _dblMinDbhSteepSlope,
                    _dblMaxCableYardingDistance, _dblMaxHelicopterCableYardingDistance, _strHarvestMethodSteepSlope, _strHarvestMethodSteepSlope,
                    _intGenericMerchAsPercentOfTotalVol);
            }
        }

        private class prescription
        {
            string _strRx = "";
            string _strHarvestMethodLowSlope = "";
            string _strHarvestMethodSteepSlope = "";
            int _intHarvestCategoryLowSlope;
            int _intHarvestCategorySteepSlope;

            public prescription(string rx, string harvestMethodLowSlope, string harvestMethodSteepSlope,
                                int harvestCategoryLowSlope, int harvestCategorySteepSlope)
            {
                _strRx = rx;
                _strHarvestMethodLowSlope = harvestMethodLowSlope;
                _strHarvestMethodSteepSlope = harvestMethodSteepSlope;
                _intHarvestCategoryLowSlope = harvestCategoryLowSlope;
                _intHarvestCategorySteepSlope = harvestCategorySteepSlope;
            }

            public string Rx
            {
                get { return _strRx; }
            }
            public string HarvestMethodLowSlope
            {
                get { return _strHarvestMethodLowSlope; }
            }
            public string HarvestMethodSteepSlope
            {
                get { return _strHarvestMethodSteepSlope; }
            }
            public int HarvestCategoryLowSlope
            {
                get { return _intHarvestCategoryLowSlope; }
            }
            public int HarvestCategorySteepSlope
            {
                get { return _intHarvestCategorySteepSlope; }
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
            string _strHarvestSystem;
            double _dblTotalBcTpa;
            double _dblTotalBcVolCf;
            double _dblTotalChipTpa;
            double _dblTotalChipNonMerchVolCf;
            double _dblTotalChipMerchVolCf;
            double _dblTotalChipVolCf;
            double _dblTotalChipWtGt;
            double _dblTotalChipHwdVolCf;
            double _dblTotalSmLogTpa;
            double _dblTotalSmLogNonCommVolCf;
            double _dblTotalSmLogMerchVolCf;
            double _dblTotalSmLogNonCommMerchVolCf;
            double _dblTotalSmLogCommNonMerchVolCf;
            double _dblTotalSmLogVolCf;
            double _dblTotalSmLogWtGt;
            double _dblTotalSmLogHwdVolCf;
            double _dblTotalLgLogTpa;
            double _dblTotalLgLogMerchVolCf;
            double _dblTotalLgLogNonCommMerchVolCf;
            double _dblTotalLgLogNonCommVolCf;
            double _dblTotalLgLogCommNonMerchVolCf;
            double _dblTotalLgLogVolCf;
            double _dblTotalLgLogWtGt;
            double _dblTotalLgLogHwdVolCf;


            public opcostInput(string condId, int percentSlope, string rxCycle, string rxPackage, string rx,
                               string rxYear, double yardingDistance, int elev, string harvestSystem)
            {
                _strCondId = condId;
                _intPercentSlope = percentSlope;
                _strRxCycle = rxCycle;
                _strRxPackage = rxPackage;
                _strRx = rx;
                _strRxYear = rxYear;
                _dblYardingDistance = yardingDistance;
                _intElev = elev;
                _strHarvestSystem = harvestSystem;
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
            public string HarvestSystem
            {
                get { return _strHarvestSystem; }
            }
            public double TotalBcTpa
            {
                set { _dblTotalBcTpa = value; }
                get { return _dblTotalBcTpa; }
            }
            public double TotalBcVolCf
            {
                set { _dblTotalBcVolCf = value; }
                get { return _dblTotalBcVolCf; }
            }
            public double TotalChipTpa
            {
                set { _dblTotalChipTpa = value; }
                get { return _dblTotalChipTpa; }
            }
            public double TotalChipMerchVolCf
            {
                set { _dblTotalChipMerchVolCf = value; }
                get { return _dblTotalChipMerchVolCf; }
            }
            public double TotalChipNonMerchVolCf
            {
                set { _dblTotalChipNonMerchVolCf = value; }
                get { return _dblTotalChipNonMerchVolCf; }
            }
            public double TotalChipVolCf
            {
                set { _dblTotalChipVolCf = value; }
                get { return _dblTotalChipVolCf; }
            }
            public double TotalChipWtGt
            {
                set { _dblTotalChipWtGt = value; }
                get { return _dblTotalChipWtGt; }
            }
            public double TotalChipHwdVolCf
            {
                set { _dblTotalChipHwdVolCf = value; }
                get { return _dblTotalChipHwdVolCf; }
            }
            public double TotalSmLogTpa
            {
                set { _dblTotalSmLogTpa = value; }
                get { return _dblTotalSmLogTpa; }
            }
            public double TotalSmLogMerchVolCf
            {
                set { _dblTotalSmLogMerchVolCf = value; }
                get { return _dblTotalSmLogMerchVolCf; }
            }
            public double TotalSmLogNonCommVolCf
            {
                set { _dblTotalSmLogNonCommVolCf = value; }
                get { return _dblTotalSmLogNonCommVolCf; }
            }
            public double TotalSmLogNonCommMerchVolCf
            {
                set { _dblTotalSmLogNonCommMerchVolCf = value; }
                get { return _dblTotalSmLogNonCommMerchVolCf; }
            }
            public double TotalSmLogCommNonMerchVolCf
            {
                set { _dblTotalSmLogCommNonMerchVolCf = value; }
                get { return _dblTotalSmLogCommNonMerchVolCf; }
            }
            public double TotalSmLogVolCf
            {
                set { _dblTotalSmLogVolCf = value; }
                get { return _dblTotalSmLogVolCf; }
            }
            public double TotalSmLogWtGt
            {
                set { _dblTotalSmLogWtGt = value; }
                get { return _dblTotalSmLogWtGt; }
            }
            public double TotalSmLogHwdVolCf
            {
                set { _dblTotalSmLogHwdVolCf = value; }
                get { return _dblTotalSmLogHwdVolCf; }
            }
            public double TotalLgLogTpa
            {
                set { _dblTotalLgLogTpa = value; }
                get { return _dblTotalLgLogTpa; }
            }
            public double TotalLgLogNonCommMerchVolCf
            {
                set { _dblTotalLgLogNonCommMerchVolCf = value; }
                get { return _dblTotalLgLogNonCommMerchVolCf; }
            }
            public double TotalLgLogMerchVolCf
            {
                set { _dblTotalLgLogMerchVolCf = value; }
                get { return _dblTotalLgLogMerchVolCf; }
            }
            public double TotalLgLogNonCommVolCf
            {
                set { _dblTotalLgLogNonCommVolCf = value; }
                get { return _dblTotalLgLogNonCommVolCf; }
            }
            public double TotalLgLogCommNonMerchVolCf
            {
                set { _dblTotalLgLogCommNonMerchVolCf = value; }
                get { return _dblTotalLgLogCommNonMerchVolCf; }
            }
            public double TotalLgLogVolCf
            {
                set { _dblTotalLgLogVolCf = value; }
                get { return _dblTotalLgLogVolCf; }
            }
            public double TotalLgLogWtGt
            {
                set { _dblTotalLgLogWtGt = value; }
                get { return _dblTotalLgLogWtGt; }
            }
            public double TotalLgLogHwdVolCf
            {
                set { _dblTotalLgLogHwdVolCf = value; }
                get { return _dblTotalLgLogHwdVolCf; }
            }
            public string CondId
            {
                set { _strCondId = value; }
                get { return _strCondId; }
            }
            public string RxPackage
            {
                get { return _strRxPackage; }
            }
            public string RxCycle
            {
                get { return _strRxCycle; }
            }
            public string Rx
            {
                get { return _strRx; }
            }
        }

        /// <summary>
        /// An treeVolValInput object represents a line in the tree vol val file
        /// The metrics are aggregated by conditionId, rxCycle, speciesGroup, and diameterGroup
        /// </summary>
        private class treeVolValInput
        {
            string _strCondId = "";
            string _strRxCycle = "";
            string _strRxPackage = "";
            string _strRx = "";
            int _intSpeciesGroup;
            int _intDiamGroup;
            int _intHarvestMethodCategory;
            string _strMerchToChip;
            double _dblChipMktValPgt;
            double _dblTotalBrushCutVolumeCf;
            double _dblTotalBrushCutWtGt;
            double _dblTotalStandResidueWtGt;

            public treeVolValInput(string condId, string rxCycle, string rxPackage, string rx,
                                    int speciesGroup, int diamGroup, bool isNonCommercial,
                                    double chipMktValPgt, int harvestMethodCategory)
            {
                _strCondId = condId;
                _strRxCycle = rxCycle;
                _strRxPackage = rxPackage;
                _strRx = rx;
                _intSpeciesGroup = speciesGroup;
                _intDiamGroup = diamGroup;
                _intHarvestMethodCategory = harvestMethodCategory;
                if (isNonCommercial)
                {
                    _strMerchToChip = "Y";
                }
                else
                {
                    _strMerchToChip = "N";
                }
                _dblChipMktValPgt = chipMktValPgt;
            }
            
            public string CondId
            {
                get { return _strCondId; }
            }
            public string Rx
            {
                get { return _strRx; }
            }
            public string RxCycle
            {
                get { return _strRxCycle; }
            }
            public string RxPackage
            {
                get { return _strRxPackage; }
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
            public string MerchToChip
            {
                get { return _strMerchToChip; }
            }
            public double ChipMktValPgt
            {
                get { return _dblChipMktValPgt; }
            }
            public int HarvestMethodCategory
            {
                get { return _intHarvestMethodCategory; }
            }
            public double TotalBrushCutVolumeCf
            {
                get { return _dblTotalBrushCutVolumeCf; }
                set { _dblTotalBrushCutVolumeCf = value; }
            }
            public double TotalBrushCutWtGt
            {
                get { return _dblTotalBrushCutWtGt; }
                set { _dblTotalBrushCutWtGt = value; }
            }
            public double TotalStandResidueWtGt
            {
                get { return _dblTotalStandResidueWtGt; }
                set { _dblTotalStandResidueWtGt = value; }
            }
        }

        private class treeSpecies
        {
            string _strSpcd = "";
            int _intSpeciesGroup;
            double _dblOdWgt;
            double _dblDryToGreen;

            public treeSpecies(string spCd, int speciesGroup, double odWgt, double dryToGreen)
            {
                _strSpcd = spCd;
                _intSpeciesGroup = speciesGroup;
                _dblOdWgt = odWgt;
                _dblDryToGreen = dryToGreen;
            }

            public string Spcd
            {
                get { return _strSpcd; }
            }
            public int SpeciesGroup
            {
                get { return _intSpeciesGroup; }
            }
            public double OdWgt
            {
                get { return _dblOdWgt; }
            }
            public double DryToGreen
            {
                get { return _dblDryToGreen; }
            }
        }

        private class escalators
        {
            double _dblEnergyWoodRevCycle2;
            double _dblEnergyWoodRevCycle3;
            double _dblEnergyWoodRevCycle4;
            double _dblMerchWoodRevCycle2;
            double _dblMerchWoodRevCycle3;
            double _dblMerchWoodRevCycle4;


            public escalators(double energyWoodRevCycle2, double energyWoodRevCycle3, double energyWoodRevCycle4,
                              double merchWoodRevCycle2, double merchWoodRevCycle3, double merchWoodRevCycle4)
            {
                _dblEnergyWoodRevCycle2 = energyWoodRevCycle2;
                _dblEnergyWoodRevCycle3 = energyWoodRevCycle3;
                _dblEnergyWoodRevCycle4 = energyWoodRevCycle4;
                _dblMerchWoodRevCycle2 = merchWoodRevCycle2;
                _dblMerchWoodRevCycle3 = merchWoodRevCycle3;
                _dblMerchWoodRevCycle4 = merchWoodRevCycle4;
            }
            
            public double EnergyWoodRevCycle2
            {
                get { return _dblEnergyWoodRevCycle2; }
            }
            public double EnergyWoodRevCycle3
            {
                get { return _dblEnergyWoodRevCycle3; }
            }
            public double EnergyWoodRevCycle4
            {
                get { return _dblEnergyWoodRevCycle4; }
            }
            public double MerchWoodRevCycle2
            {
                get { return _dblMerchWoodRevCycle2; }
            }
            public double MerchWoodRevCycle3
            {
                get { return _dblMerchWoodRevCycle3; }
            }
            public double MerchWoodRevCycle4
            {
                get { return _dblMerchWoodRevCycle4; }
            }
        }

        private class harvestMethod
        {
            bool _blnSteepSlope;
            string _strMethod;
            int _intBiosumCategory;

            public harvestMethod(bool steepSlope, string method, int biosumCategory)
            {
                _blnSteepSlope = steepSlope;
                _strMethod = method;
                _intBiosumCategory = biosumCategory;
            }

            public bool SteepSlope
            {
                get { return _blnSteepSlope; }
            }
            public string Method
            {
                get { return _strMethod; }
            }
            public int BiosumCategory
            {
                get { return _intBiosumCategory; }
            }
        }

    }
}
