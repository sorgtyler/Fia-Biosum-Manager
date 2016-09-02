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
        private string m_strScenarioId = "scenario1";
        private string m_strOpcostTableName = "OPCOST_INPUT_NEW";
        private string m_strTvvTableName = "tree_vol_val_by_species_diam_groups_new";
        private string m_strDebugFile =
            "";
        private ado_data_access m_oAdo;
        private List<tree> m_trees;
        private scenarioHarvestMethod m_scenarioHarvestMethod;
        private IDictionary<string, prescription> m_prescriptions;

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
                "\\processor\\" + m_strScenarioId + "\\" + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsDbFile;

            //
            //LOAD PROJECT DATATASOURCES INFO
            //
            m_oQueries.m_oFvs.LoadDatasource = true;
            m_oQueries.m_oReference.LoadDatasource = true;
            m_oQueries.m_oProcessor.LoadDatasource = true;
            m_oQueries.LoadDatasources(true, "processor", m_strScenarioId);

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
            //Load presciptions into reference dictionary
            m_prescriptions = loadPrescriptions();
            //Load diameter groups into reference list
            List<treeDiamGroup> listDiamGroups = loadTreeDiamGroups();
            //Load species groups into reference dictionary
            IDictionary<string, treeSpecies> dictTreeSpecies = loadTreeSpecies(p_strVariant);
            //Load species diam values into reference dictionary
            IDictionary<string, speciesDiamValue> dictSpeciesDiamValues = loadSpeciesDiamValues(m_strScenarioId);
            //Load diameter variables into reference object
            m_scenarioHarvestMethod = loadScenarioHarvestMethod(m_strScenarioId);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "loadTrees: Diameter Variables in Use: " + m_scenarioHarvestMethod.ToString() + "\r\n");


            
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));
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
                        prescription currentPrescription = null;
                        m_prescriptions.TryGetValue(newTree.Rx, out currentPrescription);
                        if (currentPrescription != null)
                        {
                            strDefaultHarvestMethodLowSlope = currentPrescription.HarvestMethodLowSlope;
                            strDefaultHarvestMethodSteepSlope = currentPrescription.HarvestMethodSteepSlope;
                        }

                        if (newTree.Slope < m_scenarioHarvestMethod.SteepSlopePct)
                        {
                            // assign low slope harvest method
                            if (! String.IsNullOrEmpty(m_scenarioHarvestMethod.HarvestMethodLowSlope))
                            {
                                newTree.HarvestMethod = m_scenarioHarvestMethod.HarvestMethodLowSlope;
                            }
                            else
                            {
                                newTree.HarvestMethod = strDefaultHarvestMethodLowSlope;
                            }
                        }
                        else
                        {
                            // assign steep slope harvest method
                            if (!String.IsNullOrEmpty(m_scenarioHarvestMethod.HarvestMethodSteepSlope))
                            {
                                newTree.HarvestMethod = m_scenarioHarvestMethod.HarvestMethodSteepSlope;
                            }
                            else
                            {
                                newTree.HarvestMethod = strDefaultHarvestMethodSteepSlope;
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


                        //if (nextTree.DiamGroup < 1)
                        //{
                        //    System.Windows.MessageBox.Show("missing diam group");
                        //}
                    }
                }
                System.Windows.MessageBox.Show(m_trees.Count + " trees");
            }
            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;
        }

        private void createOpcostInput()
        {
            int intHwdSpeciesCodeThreshold = 299; // Species codes greater than this are hardwoods
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
                            nextInput.TotalBcVolCf = nextInput.TotalBcVolCf + nextTree.OpCostBrushCutVolCf;
                        }
                        // Metrics for chip trees
                        else if (nextTree.TreeType == OpCostTreeType.CT)
                        {
                            nextInput.TotalCtTpa = nextInput.TotalCtTpa + nextTree.Tpa;
                            nextInput.TotalCtTreeBiomass = nextInput.TotalCtTreeBiomass + nextTree.DryBiot;
                            nextInput.TotalCtBoleBiomass = nextInput.TotalCtBoleBiomass + nextTree.DryBiom;
                            nextInput.TotalCtVolCf = nextInput.TotalCtVolCf + nextTree.OpCostChipVolCf;
                            nextInput.TotalCtWtGt = nextInput.TotalCtWtGt + nextTree.OpCostChipWtGt;
                            if (Convert.ToInt32(nextTree.SpCd) > intHwdSpeciesCodeThreshold)
                                nextInput.TotalCtHwdVolCf = nextInput.TotalCtHwdVolCf + nextTree.OpCostChipVolCf;

                        }
                        // Metrics for small log trees
                        else if (nextTree.TreeType == OpCostTreeType.SL)
                        {
                            nextInput.TotalSmLogTpa = nextInput.TotalSmLogTpa + nextTree.Tpa;
                            nextInput.TotalSmLogTreeBiomass = nextInput.TotalSmLogTreeBiomass + nextTree.DryBiot;
                            nextInput.TotalSmLogBoleBiomass = nextInput.TotalSmLogBoleBiomass + nextTree.DryBiom;
                            nextInput.TotalSmLogVolCf = nextInput.TotalSmLogVolCf + nextTree.OpCostMerchVolCf;
                            nextInput.TotalSmLogWtGt = nextInput.TotalSmLogWtGt + nextTree.OpCostMerchWtGt;
                            if (Convert.ToInt32(nextTree.SpCd) > intHwdSpeciesCodeThreshold)
                                nextInput.TotalSmLogHwdVolCf = nextInput.TotalSmLogHwdVolCf + nextTree.OpCostMerchVolCf;
                        }
                        // Metrics for small log trees
                        else if (nextTree.TreeType == OpCostTreeType.LL)
                        {
                            nextInput.TotalLgLogTpa = nextInput.TotalLgLogTpa + nextTree.Tpa;
                            nextInput.TotalLgLogTreeBiomass = nextInput.TotalLgLogTreeBiomass + nextTree.DryBiot;
                            nextInput.TotalLgLogBoleBiomass = nextInput.TotalLgLogBoleBiomass + nextTree.DryBiom;
                            nextInput.TotalLgLogVolCf = nextInput.TotalLgLogVolCf + nextTree.OpCostMerchVolCf;
                            nextInput.TotalLgLogWtGt = nextInput.TotalLgLogWtGt + nextTree.OpCostMerchWtGt;
                            if (Convert.ToInt32(nextTree.SpCd) > intHwdSpeciesCodeThreshold)
                                nextInput.TotalLgLogHwdVolCf = nextInput.TotalLgLogHwdVolCf + nextTree.OpCostMerchVolCf;
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
                    if (nextStand.TotalCtBoleBiomass > 0)
                        { dblCtResidueFraction = Math.Round((nextStand.TotalCtTreeBiomass - nextStand.TotalCtBoleBiomass) / nextStand.TotalCtBoleBiomass) * 100; }
                    double dblCtAvgVolume = 0;
                    if (nextStand.TotalCtTpa > 0)
                        { dblCtAvgVolume = nextStand.TotalCtVolCf / nextStand.TotalCtTpa; }
                    double dblCtAvgDensity = 0;
                    double dblCtHwdProp = 0;
                    if (nextStand.TotalCtVolCf > 0)
                    {
                        dblCtAvgDensity = nextStand.TotalCtWtGt * 2000 / nextStand.TotalCtVolCf;
                        dblCtHwdProp = nextStand.TotalCtHwdVolCf / nextStand.TotalCtVolCf;
                    }

                    
                    // *** SMALL LOGS ***
                    double dblSmLogResidueFraction = 0;
                    if (nextStand.TotalSmLogBoleBiomass > 0)
                        { dblSmLogResidueFraction = Math.Round((nextStand.TotalSmLogTreeBiomass - nextStand.TotalSmLogBoleBiomass) / nextStand.TotalSmLogBoleBiomass) * 100; }
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
                    if (nextStand.TotalLgLogBoleBiomass > 0)
                    { dblLgLogResidueFraction = Math.Round((nextStand.TotalLgLogTreeBiomass - nextStand.TotalLgLogBoleBiomass) / nextStand.TotalLgLogBoleBiomass) * 100; }
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
                    nextStand.Elev + ", '" + nextStand.HarvestSystem + "', " + nextStand.TotalCtTpa + ", " + 
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

        private void updateTreeVolVal(string strDateTimeCreated)
        {
            if (m_trees.Count < 1)
            {
                System.Windows.MessageBox.Show("No cut trees have been loaded for this scenario, variant, package combination. \r\n The tree vol val cannot be created",
                    "FIA Biosum", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return;
            }

            // load escalators; We will need them later
            escalators p_escalators = loadEscalators(m_strScenarioId);
            
            // create connection to database
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));

            if (m_oAdo.m_intError == 0)
            {                
                // create tree vol val table
                frmMain.g_oTables.m_oProcessor.CreateTreeVolValSpeciesDiamGroupsTable(m_oAdo, m_oAdo.m_OleDbConnection, m_strTvvTableName);

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createTreeVolVal: Read trees into tree vol val - " + System.DateTime.Now.ToString() + "\r\n");

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
                                chipMktValPgt = chipMktValPgt * p_escalators.EnergyWoodRevCycle2;
                                break;
                            case "3":
                                chipMktValPgt = chipMktValPgt * p_escalators.EnergyWoodRevCycle3;
                                break;
                            case "4":
                                chipMktValPgt = chipMktValPgt * p_escalators.EnergyWoodRevCycle4;
                                break;
                        }
                        nextInput = new treeVolValInput(nextTree.CondId, nextTree.RxCycle, nextTree.RxPackage, nextTree.Rx,
                            nextTree.SpeciesGroup, nextTree.DiamGroup, nextTree.IsNonCommercial, chipMktValPgt);
                        dictTvvInput.Add(strKey, nextInput);
                    }
                }
                
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createTreeVolVal: Finished reading trees - " + System.DateTime.Now.ToString() + "\r\n");

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createTreeVolVal: Begin writing tree vol val table - " + System.DateTime.Now.ToString() + "\r\n");
                long lngCount =0;
                foreach (string key in dictTvvInput.Keys)
                {
                    treeVolValInput nextStand = dictTvvInput[key];
                    m_oAdo.m_strSQL = "INSERT INTO " + m_strTvvTableName + " " +
                    "(biosum_cond_id, rxpackage, rx, rxcycle, species_group, diam_group, " +
                    "merch_to_chipbin_YN, chip_mkt_val_pgt, DateTimeCreated, place_holder)" +
                    "VALUES ('" + nextStand.CondId + "', '" + nextStand.RxPackage + "', '" + nextStand.Rx + "', '" + 
                    nextStand.RxCycle + "', " + nextStand.SpeciesGroup + ", " + nextStand.DiamGroup +
                    ", '" + nextStand.MerchToChip + "', " + nextStand.ChipMktValPgt + ", '" + strDateTimeCreated + "', 'N')";

                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    if (m_oAdo.m_intError != 0) break;
                    lngCount++;

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "END createTreeVolVal INSERTED " + lngCount + " RECORDS: " + System.DateTime.Now.ToString() + "\r\n");
                }
            }
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

        private IDictionary<String, treeSpecies> loadTreeSpecies(string p_strVariant)
        {
            IDictionary<String, treeSpecies> dictTreeSpecies = new Dictionary<String, treeSpecies>();
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT DISTINCT SPCD, USER_SPC_GROUP, OD_WGT, Dry_to_Green FROM tree_species " +
                                "WHERE FVS_VARIANT = '" + p_strVariant + "' " +
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

        private IDictionary<String, prescription> loadPrescriptions()
        {
            IDictionary<String, prescription> dictPrescriptions = new Dictionary<String, prescription>();
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT * FROM rx";
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        string strRx = Convert.ToString(m_oAdo.m_OleDbDataReader["rx"]).Trim();
                        string strHarvestMethodLowSlope = Convert.ToString(m_oAdo.m_OleDbDataReader["HarvestMethodLowSlope"]).Trim();
                        string strHarvestMethodSteepSlope = Convert.ToString(m_oAdo.m_OleDbDataReader["HarvestMethodSteepSlope"]).Trim();

                        dictPrescriptions.Add(strRx, new prescription(strRx, strHarvestMethodLowSlope, strHarvestMethodSteepSlope));
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
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));
            scenarioHarvestMethod returnVariables = null;
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
                    double dblMaxCableYardingDistance = Convert.ToDouble(m_oAdo.m_OleDbDataReader["MaxCableYardingDistance"]);
                    double dblMaxHelicopterCableYardingDistance = Convert.ToDouble(m_oAdo.m_OleDbDataReader["MaxHelicopterCableYardingDistance"]);
                    string strHarvestMethodLowSlope = Convert.ToString(m_oAdo.m_OleDbDataReader["HarvestMethodLowSlope"]).Trim();
                    string strHarvestMethodSteepSlope = Convert.ToString(m_oAdo.m_OleDbDataReader["HarvestMethodSteepSlope"]).Trim();

                    returnVariables = new scenarioHarvestMethod(dblMinChipDbh, dblMinSmallLogDbh, dblMinLgLogDbh,
                        intMinSlopePct, dblMinDbhSteepSlope, dblMaxCableYardingDistance, dblMaxHelicopterCableYardingDistance,
                        strHarvestMethodLowSlope, strHarvestMethodSteepSlope);
                }
            }

            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;

            return returnVariables;
        }

        private escalators loadEscalators(string p_scenario)
        {
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));
            escalators returnEscalators = null;
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT * FROM scenario_cost_revenue_escalators " +
                                "WHERE scenario_id = '" + p_scenario + "'";
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    // We should only have one record
                    m_oAdo.m_OleDbDataReader.Read();
                    double dblEnergyWoodRevCycle2 = Convert.ToDouble(m_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle2"]);
                    double dblEnergyWoodRevCycle3 = Convert.ToDouble(m_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle3"]);
                    double dblEnergyWoodRevCycle4 = Convert.ToDouble(m_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle4"]);

                    returnEscalators = new escalators(dblEnergyWoodRevCycle2, dblEnergyWoodRevCycle3, dblEnergyWoodRevCycle4);
                }
            }

            // Always close the connection
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo = null;

            return returnEscalators;
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
            if (p_tree.TreeType == OpCostTreeType.LL || p_tree.TreeType == OpCostTreeType.SL)
            {
                p_tree.OpCostMerchVolCf = p_tree.VolCfNet * p_tree.Tpa;
                if (p_tree.DryToGreen != 0)
                    { p_tree.OpCostMerchWtGt = p_tree.VolCfNet * p_tree.OdWgt * p_tree.Tpa / p_tree.DryToGreen / 2000; }
            }
            else if (p_tree.TreeType == OpCostTreeType.BC)
            {
                //p_tree.BrushCutWtGt = (p_tree.DryBiot * p_tree.Tpa / p_tree.DryToGreen) / 2000;
                p_tree.OpCostBrushCutVolCf = p_tree.DryBiot * p_tree.Tpa / p_tree.OdWgt;
            }
            else if (p_tree.TreeType == OpCostTreeType.CT)
            {
                // For opCost chip trees we just send bole volume
                // OpCost is expects bole volumes/weights only for chip trees.
                if (p_tree.VolCfNet > 0)
                {
                    p_tree.OpCostChipVolCf = p_tree.VolCfNet * p_tree.Tpa;
                }
                else
                {
                    p_tree.OpCostChipVolCf = p_tree.VolTsGrs * p_tree.Tpa;
                }

                // Just sending weight of bole as well for chips
                if (p_tree.VolCfNet > 0)
                {
                    p_tree.OpCostChipWtGt = (p_tree.VolCfNet * p_tree.OdWgt * p_tree.Tpa) / p_tree.DryToGreen / 2000;
                }
                else
                {
                    p_tree.OpCostChipWtGt = (p_tree.VolTsGrs * p_tree.OdWgt * p_tree.Tpa) / p_tree.DryToGreen / 2000;
                }
            }
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
            OpCostTreeType _TreeType;
            int _intSpeciesGroup;
            int _intDiamGroup;
            bool _boolIsNonCommercial;
            double _dblMerchValue;
            double _dblChipValue;
            int _intElev;
            double _dblYardingDistance;
            string _strHarvestMethod;
            double _dblOdWgt;
            double _dblDryToGreen;
            double _dblMerchVolCf;
            double _dblOpCostMerchVolCf;
            double _dblMerchWtGt;
            double _dblOpCostMerchWtGt;
            double _dblBrushCutVolCf;
            double _dblBrushCutWtGt;
            double _dblOpCostBrushCutVolCf;
            double _dblChipVolCf;
            double _dblOpCostChipVolCf;
            double _dblChipWtGt;
            double _dblOpCostChipWtGt;

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
            public double OpCostMerchVolCf
            {
                get { return _dblOpCostMerchVolCf; }
                set { _dblOpCostMerchVolCf = value; }
            }
            public double MerchWtGt
            {
                get { return _dblMerchWtGt; }
                set { _dblMerchWtGt = value; }
            }
            public double OpCostMerchWtGt
            {
                get { return _dblOpCostMerchWtGt; }
                set { _dblOpCostMerchWtGt = value; }
            }
            public double BrushCutVolCf
            {
                get { return _dblBrushCutVolCf; }
                set { _dblBrushCutVolCf = value; }
            }
            public double OpCostBrushCutVolCf
            {
                get { return _dblOpCostBrushCutVolCf; }
                set { _dblOpCostBrushCutVolCf = value; }
            }
            public double BrushCutWtGt
            {
                get { return _dblBrushCutWtGt; }
                set { _dblBrushCutWtGt = value; }
            }
            public double ChipVolCf
            {
                get { return _dblChipVolCf; }
                set { _dblChipVolCf = value; }
            }
            public double OpCostChipVolCf
            {
                get { return _dblOpCostChipVolCf; }
                set { _dblOpCostChipVolCf = value; }
            }
            public double ChipWtGt
            {
                get { return _dblChipWtGt; }
                set { _dblChipWtGt = value; }
            }
            public double OpCostChipWtGt
            {
                get { return _dblOpCostChipWtGt; }
                set { _dblOpCostChipWtGt = value; }
            }
            public string HarvestMethod
            {
                get { return _strHarvestMethod; }
                set { _strHarvestMethod = value; }
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

            public scenarioHarvestMethod(double minChipDbh, double minSmallLogDbh, double minLargeLogDbh, int steepSlopePct,
                                         double minDbhSteepSlope, double maxCableYardingDistance, double maxHelicopterYardingDistance,
                                         string harvestMethodLowSlope, string harvestMethodSteepSlope)
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
            // Overriding the ToString method for debugging purposes
            public override string ToString()
            {
                return string.Format("MinChipDbh: {0}, MinSmallLogDbh: {1}, MinLargeLogDbh: {2}, SteepSlopePct: {3}, MinDbhSteepSlope: {4}, " +
                    "MaxCableYardingDistance: {5}, MaxHelicopterCableYardingDistance: {6}, HarvestMethodLowSlope: {7}, HarvestMethodSteepSlope: {8} ]",
                    _dblMinChipDbh, _dblMinSmallLogDbh, _dblMinLargeLogDbh, _intSteepSlopePct, _dblMinDbhSteepSlope,
                    _dblMaxCableYardingDistance, _dblMaxHelicopterCableYardingDistance, _strHarvestMethodSteepSlope, _strHarvestMethodSteepSlope);
            }
        }

        private class prescription
        {
            string _strRx = "";
            string _strHarvestMethodLowSlope = "";
            string _strHarvestMethodSteepSlope = "";

            public prescription(string rx, string harvestMethodLowSlope, string harvestMethodSteepSlope)
            {
                _strRx = rx;
                _strHarvestMethodLowSlope = harvestMethodLowSlope;
                _strHarvestMethodSteepSlope = harvestMethodSteepSlope;
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
            double _dblTotalCtTpa;
            double _dblTotalCtTreeBiomass;
            double _dblTotalCtBoleBiomass;
            double _dblTotalCtVolCf;
            double _dblTotalCtWtGt;
            double _dblTotalCtHwdVolCf;
            double _dblTotalSmLogTpa;
            double _dblTotalSmLogTreeBiomass;
            double _dblTotalSmLogBoleBiomass;
            double _dblTotalSmLogVolCf;
            double _dblTotalSmLogWtGt;
            double _dblTotalSmLogHwdVolCf;
            double _dblTotalLgLogTpa;
            double _dblTotalLgLogTreeBiomass;
            double _dblTotalLgLogBoleBiomass;
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
            public double TotalCtTpa
            {
                set { _dblTotalCtTpa = value; }
                get { return _dblTotalCtTpa; }
            }
            public double TotalCtBoleBiomass
            {
                set { _dblTotalCtBoleBiomass = value; }
                get { return _dblTotalCtBoleBiomass; }
            }
            public double TotalCtTreeBiomass
            {
                set { _dblTotalCtTreeBiomass = value; }
                get { return _dblTotalCtTreeBiomass; }
            }
            public double TotalCtVolCf
            {
                set { _dblTotalCtVolCf = value; }
                get { return _dblTotalCtVolCf; }
            }
            public double TotalCtWtGt
            {
                set { _dblTotalCtWtGt = value; }
                get { return _dblTotalCtWtGt; }
            }
            public double TotalCtHwdVolCf
            {
                set { _dblTotalCtHwdVolCf = value; }
                get { return _dblTotalCtHwdVolCf; }
            }
            public double TotalSmLogTpa
            {
                set { _dblTotalSmLogTpa = value; }
                get { return _dblTotalSmLogTpa; }
            }
            public double TotalSmLogBoleBiomass
            {
                set { _dblTotalSmLogBoleBiomass = value; }
                get { return _dblTotalSmLogBoleBiomass; }
            }
            public double TotalSmLogTreeBiomass
            {
                set { _dblTotalSmLogTreeBiomass = value; }
                get { return _dblTotalSmLogTreeBiomass; }
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
            public double TotalLgLogBoleBiomass
            {
                set { _dblTotalLgLogBoleBiomass = value; }
                get { return _dblTotalLgLogBoleBiomass; }
            }
            public double TotalLgLogTreeBiomass
            {
                set { _dblTotalLgLogTreeBiomass = value; }
                get { return _dblTotalLgLogTreeBiomass; }
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
            string _strMerchToChip;
            double _dblChipMktValPgt;

            public treeVolValInput(string condId, string rxCycle, string rxPackage, string rx,
                                    int speciesGroup, int diamGroup, bool isNonCommercial,
                                    double chipMktValPgt)
            {
                _strCondId = condId;
                _strRxCycle = rxCycle;
                _strRxPackage = rxPackage;
                _strRx = rx;
                _intSpeciesGroup = speciesGroup;
                _intDiamGroup = diamGroup;
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

            public escalators(double energyWoodRevCycle2, double energyWoodRevCycle3, double energyWoodRevCycle4)
            {
                _dblEnergyWoodRevCycle2 = energyWoodRevCycle2;
                _dblEnergyWoodRevCycle3 = energyWoodRevCycle3;
                _dblEnergyWoodRevCycle4 = energyWoodRevCycle4;
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
        }

    }
}
