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
        private string m_strDebugFile = frmMain.g_oEnv.strTempDir + "\\biosum_fvs_input_debug.txt";
        private ado_data_access m_oAdo;
        private List<tree> m_trees;

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
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));
            string strTableName = "fvs_tree_IN_" + p_strVariant + "_P" + p_strRxPackage + "_TREE_CUTLIST";
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT z.biosum_cond_id, z.rxCycle, z.rx, z.rxYear, " +
                                "z.dbh, z.tpa, z.volCfNet, z.drybiot, z.drybiom,z.FvsCreatedTree_YN, " +
                                "z.fvs_tree_id, z.fvs_species, " +
                                "c.slope " +
                                  "FROM " + strTableName + " z, cond c " +
                                  "WHERE z.rxpackage='" + p_strRxPackage + "' AND " +
                                  "z.biosum_cond_id = c.biosum_cond_id AND " +
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
                        m_trees.Add(newTree);
                    }
                }

                //Query TREE table to get original FIA species codes
                //@ToDo: For now loading the whole TREE table since I can't think of a more efficient way short of asking
                //for all the condition ID's
                strSQL = "SELECT t.fvs_tree_id, t.spcd " +
                        "FROM tree t " +
                        "WHERE mid(t.fvs_tree_id,1,2)='" + p_strVariant + "' ";
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

                    foreach (tree nextTree in m_trees)
                    {
                        if (!nextTree.FvsCreatedTree)
                        {
                            nextTree.SpCd = dictSpCd[nextTree.FvsTreeId];
                        }
                    }

                }
                
                //@ToDo: Don't forget to close the connection!!!
                System.Windows.MessageBox.Show(m_trees.Count + " trees");
            }

        }

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
        }
    }
}
