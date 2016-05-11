using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
    public partial class uc_scenario_processor_scenario_select : UserControl
    {
        private ProcessorScenarioItem_Collection m_oProcessorScenarioItem_Collection = new ProcessorScenarioItem_Collection();
        public ProcessorScenarioItem m_oProcessorScenarioItem;
        private Queries m_oQueries = new Queries();
        private ProcessorScenarioTools m_oProcessorScenarioTools = new ProcessorScenarioTools();
        private ListViewAlternateBackgroundColors m_oLvAlternateColors = new ListViewAlternateBackgroundColors();

        const int COL_CHECKBOX = 0;
        const int COL_SCENARIOID = 1;
        const int COL_DESC = 2;
        private bool m_bSuppressCheckEvents = false;
        
        FIA_Biosum_Manager.frmCoreScenario _frmScenario = null;
        public uc_scenario_processor_scenario_select()
        {
            InitializeComponent();
        }
        public FIA_Biosum_Manager.frmCoreScenario ReferenceCoreScenarioForm
        {
            get { return _frmScenario; }
            set { _frmScenario = value; }
        }
        public void loadvalues()
        {
            int x;
            ProcessorScenarioTools oTools = new ProcessorScenarioTools();
            lvProcessorScenario.Items.Clear();
            System.Windows.Forms.ListViewItem entryListItem = null;
            this.m_oLvAlternateColors.InitializeRowCollection();
            this.m_oLvAlternateColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
            this.m_oLvAlternateColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceForegroundColor = frmMain.g_oGridViewRowForegroundColor;
            this.m_oLvAlternateColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceListView = this.lvProcessorScenario;
            this.m_oLvAlternateColors.CustomFullRowSelect = true;
            if (frmMain.g_oGridViewFont != null) this.lvProcessorScenario.Font = frmMain.g_oGridViewFont;
            //
            //OPEN CONNECTION TO DB FILE CONTAINING PROCESSOR SCENARIO TABLE
            //
            //scenario mdb connection
            string strProcessorScenarioMDB =
              frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
              "\\processor\\db\\scenario_processor_rule_definitions.mdb";
            //
            //get a list of all the scenarios
            //
            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(strProcessorScenarioMDB,"",""));
            string[] strScenarioArray = 
                frmMain.g_oUtils.ConvertListToArray(
                    oAdo.CreateCommaDelimitedList(
                       oAdo.m_OleDbConnection,
                       "SELECT scenario_id " + 
                       "FROM scenario " + 
                       "WHERE scenario_id IS NOT NULL AND " + 
                                         "LEN(TRIM(scenario_id)) > 0",""),",");
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            if (strScenarioArray == null) return;

            for (x = 0; x <= strScenarioArray.Length - 1; x++)
            {
                //
                //LOAD PROJECT DATATASOURCES INFO
                //
                m_oQueries.m_oFvs.LoadDatasource = true;
                m_oQueries.m_oFIAPlot.LoadDatasource = true;
                m_oQueries.m_oProcessor.LoadDatasource = true;
                m_oQueries.m_oReference.LoadDatasource = true;
                m_oQueries.LoadDatasources(true, "processor", strScenarioArray[x]);
                m_oQueries.m_oDataSource.CreateScenarioRuleDefinitionTableLinks(
                    m_oQueries.m_strTempDbFile,
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim(),
                    "P");
                oTools.LoadAll(m_oQueries.m_strTempDbFile, m_oQueries,strScenarioArray[x], m_oProcessorScenarioItem_Collection);
                //System.IO.File.Delete(m_oQueries.m_strTempDbFile);

            }
            for (x = 0; x <= m_oProcessorScenarioItem_Collection.Count - 1; x++)
            {
                entryListItem = lvProcessorScenario.Items.Add(" ");

                entryListItem.UseItemStyleForSubItems = false;
                this.m_oLvAlternateColors.AddRow();
                this.m_oLvAlternateColors.AddColumns(x, lvProcessorScenario.Columns.Count);
                

                entryListItem.SubItems.Add(m_oProcessorScenarioItem_Collection.Item(x).ScenarioId);
                entryListItem.SubItems.Add(m_oProcessorScenarioItem_Collection.Item(x).Description);
                


            }
            this.m_oLvAlternateColors.ListView();

            string strScenarioMDB =
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\core\\db\\scenario_core_rule_definitions.mdb";


            string strConn = oAdo.getMDBConnString(strScenarioMDB, "", "");
            oAdo.OpenConnection(strConn);
            
            if (oAdo.TableExist(oAdo.m_OleDbConnection, Tables.CoreScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName))
            {
                oAdo.m_strSQL = "SELECT * FROM " + Tables.CoreScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName + " " +
                                "WHERE TRIM(UCASE(scenario_id)) = '" +
                                       ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";
                oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                string strProcessorScenario = "";
                string strFullDetailsYN = "N";
                
                if (oAdo.m_OleDbDataReader.HasRows)
                {
                    while (oAdo.m_OleDbDataReader.Read())
                    {
                        if (oAdo.m_OleDbDataReader["processor_scenario_id"] != System.DBNull.Value &&
                            oAdo.m_OleDbDataReader["processor_scenario_id"].ToString().Trim().Length > 0)
                        {
                            strProcessorScenario = oAdo.m_OleDbDataReader["processor_scenario_id"].ToString().Trim();
                        }
                        if (oAdo.m_OleDbDataReader["FullDetailsYN"] != System.DBNull.Value &&
                            oAdo.m_OleDbDataReader["FullDetailsYN"].ToString().Trim().Length > 0)
                        {
                            strFullDetailsYN = oAdo.m_OleDbDataReader["FullDetailsYN"].ToString().Trim();
                        }
                    }
                }
                oAdo.m_OleDbDataReader.Close();
                if (lvProcessorScenario.Items.Count > 0)
                {
                    for (x = 0; x <= lvProcessorScenario.Items.Count - 1; x++)
                    {
                        if (lvProcessorScenario.Items[x].SubItems[COL_SCENARIOID].Text.Trim().ToUpper() ==
                            strProcessorScenario.ToUpper())
                        {
                            lvProcessorScenario.Items[x].Checked = true;
                            for (int y = 0; y <= ReferenceCoreScenarioForm.uc_scenario_processor_scenario_select1.m_oProcessorScenarioItem_Collection.Count - 1; y++)
                            {
                                if (lvProcessorScenario.Items[x].SubItems[COL_SCENARIOID].Text.Trim().ToUpper() ==
                                    ReferenceCoreScenarioForm.uc_scenario_processor_scenario_select1.m_oProcessorScenarioItem_Collection.Item(y).ScenarioId.Trim().ToUpper())
                                {
                                    ReferenceCoreScenarioForm.uc_scenario_cond_filter1.strLowYardingDistanceDefault =
                                      Convert.ToString(ReferenceCoreScenarioForm.uc_scenario_processor_scenario_select1.m_oProcessorScenarioItem_Collection.Item(y).m_oHarvestMethod.MaxCableYardingDistance);

                                    ReferenceCoreScenarioForm.uc_scenario_cond_filter1.strSteepYardingDistanceDefault =
                                        Convert.ToString(ReferenceCoreScenarioForm.uc_scenario_processor_scenario_select1.m_oProcessorScenarioItem_Collection.Item(y).m_oHarvestMethod.MaxHelicopterCableYardingDistance);

                                    ReferenceCoreScenarioForm.uc_scenario_cond_filter1.strLowSlope =
                                        ReferenceCoreScenarioForm.uc_scenario_processor_scenario_select1.m_oProcessorScenarioItem_Collection.Item(y).m_oHarvestMethod.SteepSlopePercent;

                                    ReferenceCoreScenarioForm.uc_scenario_cond_filter1.strSteepSlope =
                                        ReferenceCoreScenarioForm.uc_scenario_processor_scenario_select1.m_oProcessorScenarioItem_Collection.Item(y).m_oHarvestMethod.SteepSlopePercent;
                                }
                            }
                            break;
                        }
                    }
                    if (x <= lvProcessorScenario.Items.Count - 1) lvProcessorScenario.Items[0].Selected = true;
                }

                if (strFullDetailsYN == "Y")
                    chkFullDetails.Checked = true;
                else
                    chkFullDetails.Checked = false;
            }
            else
            {
                frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioProcessorScenarioSelectTable(oAdo, oAdo.m_OleDbConnection, Tables.CoreScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName);
            }
            oAdo.CloseConnection(oAdo.m_OleDbConnection);

            

           


            


        }
        public void savevalues()
        {
            
            ado_data_access oAdo = new ado_data_access();
			string strScenarioId = this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
			string strScenarioMDB = 
				frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + 
				"\\core\\db\\scenario_core_rule_definitions.mdb";
			oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioMDB,"",""));
            oAdo.m_strSQL = "DELETE FROM " + Tables.CoreScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName + " " +
                                "WHERE TRIM(UCASE(scenario_id)) = '" +
                                       ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            if (lvProcessorScenario.CheckedItems.Count > 0)
            {
                string strColumnsList = "scenario_id,processor_scenario_id,FullDetailsYN";
                string strValuesList = "";
                strValuesList = "'" + ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim() + "',";
                strValuesList = strValuesList + "'" + lvProcessorScenario.CheckedItems[0].SubItems[COL_SCENARIOID].Text.Trim() + "',";
                if (this.chkFullDetails.Checked)
                    strValuesList = strValuesList + "'Y'";
                else
                    strValuesList = strValuesList + "'N'";

                oAdo.m_strSQL = "INSERT INTO " + Tables.CoreScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName + " " +
                                "(" + strColumnsList + ") " +
                                "VALUES " +
                                "(" + strValuesList + ")";

                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            }
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
        }

        private void lvProcessorScenario_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvProcessorScenario.SelectedItems.Count > 0)
            {
                m_oLvAlternateColors.DelegateListViewItem(lvProcessorScenario.SelectedItems[0]);
                if (chkFullDetails.Checked)
                {
                    FullDetails();
                }
                else
                {
                    txtDetails.Text = lvProcessorScenario.SelectedItems[0].SubItems[COL_DESC].Text;
                }
            }
        }

        private void lvProcessorScenario_MouseUp(object sender, MouseEventArgs e)
        {
            
            try
            {
                if (e.Button == MouseButtons.Left)
                {
                    int intRowHt = lvProcessorScenario.Items[0].Bounds.Height;
                    double dblRow = (double)(e.Y / intRowHt);
                    lvProcessorScenario.Items[lvProcessorScenario.TopItem.Index + (int)dblRow - 1].Selected = true;
                    this.m_oLvAlternateColors.DelegateListViewItem(lvProcessorScenario.Items[lvProcessorScenario.TopItem.Index + (int)dblRow - 1]);


                }
            }
            catch
            {
            }
        }

        private void lvProcessorScenario_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            
        }

        private void lvProcessorScenario_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (m_bSuppressCheckEvents == true) return;
            if (this.lvProcessorScenario.SelectedItems.Count == 0) return;
            m_bSuppressCheckEvents = true;
            for (int x = 0; x <= this.lvProcessorScenario.Items.Count - 1; x++)
            {
                if (this.lvProcessorScenario.Items[x].Index !=
                    this.lvProcessorScenario.SelectedItems[0].Index)
                {
                    lvProcessorScenario.Items[x].Checked = false;
                }
                else
                {
                    if (e.NewValue == System.Windows.Forms.CheckState.Checked)
                    {

                        ReferenceCoreScenarioForm.uc_scenario_cond_filter1.strLowYardingDistanceDefault =
                            Convert.ToString(m_oProcessorScenarioItem_Collection.Item(lvProcessorScenario.SelectedItems[0].Index).m_oHarvestMethod.MaxCableYardingDistance);

                        ReferenceCoreScenarioForm.uc_scenario_cond_filter1.strSteepYardingDistanceDefault =
                            Convert.ToString(m_oProcessorScenarioItem_Collection.Item(lvProcessorScenario.SelectedItems[0].Index).m_oHarvestMethod.MaxHelicopterCableYardingDistance);

                        ReferenceCoreScenarioForm.uc_scenario_cond_filter1.strLowSlope =
                            m_oProcessorScenarioItem_Collection.Item(lvProcessorScenario.SelectedItems[0].Index).m_oHarvestMethod.SteepSlopePercent;

                        ReferenceCoreScenarioForm.uc_scenario_cond_filter1.strSteepSlope =
                            m_oProcessorScenarioItem_Collection.Item(lvProcessorScenario.SelectedItems[0].Index).m_oHarvestMethod.SteepSlopePercent;

                    }
                }

            }
            m_bSuppressCheckEvents = false;

        }

        private void panel1_Resize(object sender, EventArgs e)
        {

            this.txtDetails.Width = this.panel1.Width - (int)(txtDetails.Left * 2);
            this.lvProcessorScenario.Width = txtDetails.Width;
        }

        private void chkFullDetails_CheckedChanged(object sender, EventArgs e)
        {
            if (chkFullDetails.Checked && lvProcessorScenario.SelectedItems.Count > 0)
                   FullDetails();
        }
        private void FullDetails()
        {
            
            this.txtDetails.Text  = "";
            this.m_oProcessorScenarioItem = m_oProcessorScenarioItem_Collection.Item(lvProcessorScenario.SelectedItems[0].Index);
            this.txtDetails.Text = m_oProcessorScenarioTools.ScenarioProperties(m_oProcessorScenarioItem);



        }
        public int val_processorscenario()
        {
            if (this.lvProcessorScenario.Items.Count == 0)
            {
                MessageBox.Show("Run Scenario Failed: No Processor Scenarios exist. At least one Processor Scenario must exist to run a Core Analysis Scenario", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return -1;
            }
            if (this.lvProcessorScenario.CheckedItems.Count == 0)
            {
                MessageBox.Show("Run Scenario Failed: Select at least one processor scenario in <Cost and Revenue><Processor Scenario>", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return -1;
            }
            this.m_oProcessorScenarioItem = this.m_oProcessorScenarioItem_Collection.Item(lvProcessorScenario.CheckedItems[0].Index);

            return 0;

        }
    }
}
