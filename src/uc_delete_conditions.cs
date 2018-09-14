using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity.ModelConfiguration.Conventions;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Devart.Common;

namespace FIA_Biosum_Manager
{
    public partial class uc_delete_conditions : UserControl
    {
        public int m_DialogHt;
        public int m_DialogWd;
        private DataTable m_dtStateCounty;
        private DataTable m_dtPlot;
        private Datasource m_oDatasource;
        private string m_strPlotTable;
        private string m_strCondTable;
        private string m_strTreeTable;
        private string m_strSiteTreeTable;
        private string m_strTreeRegionalBiomassTable;
        private string m_strPpsaTable;
        private string m_strPopEstUnitTable;
        private string m_strPopStratumTable;
        private string m_strPopEvalTable;
        private string m_strBiosumPopStratumAdjustmentFactorsTable;
        private string m_strTreeMacroPlotBreakPointDiaTable;

        private HashSet<string> setBiosumCondCNs = new HashSet<string>();
        private HashSet<string> setBiosumCondIds = new HashSet<string>();
        private HashSet<string> setBiosumPlotIds = new HashSet<string>();
        private HashSet<string> setPlotCNs = new HashSet<string>();
        private HashSet<string> setTreeCNs = new HashSet<string>();
        private string m_strCondCNs;
        private string m_strBiosumCondCNs;
        private string m_strBiosumCondIds;
        private string m_strBiosumPlotIds;
        private string m_strPlotCNs;
        private string m_strTreeCNs;
        Dictionary<string, object[]> m_dictIdentityColumnsToValues;
        Dictionary<string, Dictionary<string, int>> m_dictDeletedRowCountsByDatabaseAndTable =
            new Dictionary<string, Dictionary<string, int>>();

        //TODO: Help files
        private env m_oEnv;

        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultDatabaseXPSFile;
        private int m_intError;
        private string m_strStateCountyPlotSQL;
        private string m_strStateCountySQL;
        private string m_strCurrentProcess;
        private frmTherm m_frmTherm;
        private ado_data_access m_ado;
        private dao_data_access m_dao;
        private OleDbConnection m_connTempMDBFile;
        private string m_strTempMDBFileConn;
        private string m_strTempMDBFile;
        private string m_strSQL;
        private string m_strProjDir;
        private string m_strMessage = "";

        public uc_delete_conditions()
        {
            ReferenceFormDialog = null;
            InitializeComponent();
            Initialize();
        }

        private void Initialize()
        {
            Width = 700;
            m_DialogWd = Width + 10;
            m_DialogHt = groupBox1.Top + grpboxFilter.Top + grpboxFilter.Height + 100;

            grpboxFilterByState.Left = grpboxFilter.Left;
            grpboxFilterByState.Width = grpboxFilter.Width;
            grpboxFilterByState.Height = grpboxFilter.Height;
            grpboxFilterByState.Top = grpboxFilter.Top;
            btnFilterByStateHelp.Location = btnFilterHelp.Location;
            btnFilterByStateCancel.Location = btnFilterCancel.Location;
            btnFilterByStatePrevious.Location = btnFilterPrevious.Location;
            btnFilterByStateNext.Location = btnFilterNext.Location;
            btnFilterByStateFinish.Location = btnFilterFinish.Location;
            grpboxFilterByState.Visible = false;

            grpboxFilterByCondId.Left = grpboxFilter.Left;
            grpboxFilterByCondId.Width = grpboxFilter.Width;
            grpboxFilterByCondId.Height = grpboxFilter.Height;
            grpboxFilterByCondId.Top = grpboxFilter.Top;
            btnFilterByPlotHelp.Location = btnFilterHelp.Location;
            btnFilterByPlotCancel.Location = btnFilterCancel.Location;
            btnFilterByPlotPrevious.Location = btnFilterPrevious.Location;
            btnFilterByPlotNext.Location = btnFilterNext.Location;
            btnFilterByPlotFinish.Location = btnFilterFinish.Location;
            grpboxFilterByCondId.Visible = false;

            lstFilterByState.Clear();
            lstFilterByState.Columns.Add(" ", 100, HorizontalAlignment.Center);
            lstFilterByState.Columns.Add("State", 100, HorizontalAlignment.Left);
            lstFilterByState.Columns.Add("County", 100, HorizontalAlignment.Left);

            //create state,count table
            m_dtStateCounty = new DataTable("statecounty");
            m_dtStateCounty.Columns.Add("statecd", typeof(string));
            m_dtStateCounty.Columns.Add("countycd", typeof(string));

            // two columns in the Primary Key.
            var colPk = new DataColumn[2];
            colPk[0] = m_dtStateCounty.Columns["statecd"];
            colPk[1] = m_dtStateCounty.Columns["countycd"];
            m_dtStateCounty.PrimaryKey = colPk;

            //create state,county,plot table
            m_dtPlot = new DataTable("statecountyplot");
            m_dtPlot.Columns.Add("statecd", typeof(string));
            m_dtPlot.Columns.Add("countycd", typeof(string));
            m_dtPlot.Columns.Add("plot", typeof(string));

            rdoFilterByMenu.Enabled = false;
            rdoDeleteAllConds.Enabled = false;

            m_oEnv = new env();
        }

        private void InitializeDatasource()
        {
            m_strProjDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();

            m_oDatasource = new Datasource();
            m_oDatasource.LoadTableColumnNamesAndDataTypes = false;
            m_oDatasource.LoadTableRecordCount = false;
            m_oDatasource.m_strDataSourceMDBFile = m_strProjDir + "\\db\\project.mdb";
            m_oDatasource.m_strDataSourceTableName = "datasource";
            m_oDatasource.m_strScenarioId = "";
            m_oDatasource.populate_datasource_array();

            //get table names
            m_strPlotTable = m_oDatasource.getValidDataSourceTableName("PLOT");
            m_strCondTable = m_oDatasource.getValidDataSourceTableName("CONDITION");
            m_strTreeTable = m_oDatasource.getValidDataSourceTableName("TREE");
            m_strSiteTreeTable = m_oDatasource.getValidDataSourceTableName("SITE TREE");
            m_strTreeRegionalBiomassTable = m_oDatasource.getValidDataSourceTableName("TREE REGIONAL BIOMASS");
            m_strPpsaTable = m_oDatasource.getValidDataSourceTableName("POPULATION PLOT STRATUM ASSIGNMENT");
            m_strPopEstUnitTable = m_oDatasource.getValidDataSourceTableName("POPULATION ESTIMATION UNIT");
            m_strPopStratumTable = m_oDatasource.getValidDataSourceTableName("POPULATION STRATUM");
            m_strPopEvalTable = m_oDatasource.getValidDataSourceTableName("POPULATION EVALUATION");
            m_strBiosumPopStratumAdjustmentFactorsTable =
                m_oDatasource.getValidDataSourceTableName("BIOSUM POP STRATUM ADJUSTMENT FACTORS");
            m_strTreeMacroPlotBreakPointDiaTable =
                m_oDatasource.getValidDataSourceTableName("FIA TREE MACRO PLOT BREAKPOINT DIAMETER");
        }

        private void rdoFilterByFile_Click(object sender, System.EventArgs e)
        {
            this.btnFilterFinish.Enabled = false;
            this.btnFilterNext.Enabled = false;
            this.txtFilterByFile.Enabled = true;
            this.btnFilterByFileBrowse.Enabled = true;
            if (!string.IsNullOrWhiteSpace(txtFilterByFile.Text))
            {
                this.btnFilterFinish.Enabled = true;
            }
        }

        private void rdoFilterByMenu_Click(object sender, System.EventArgs e)
        {
            this.btnFilterFinish.Enabled = false;
            this.btnFilterNext.Enabled = true;
            this.txtFilterByFile.Enabled = false;
            this.btnFilterByFileBrowse.Enabled = false;
        }

        private void rdoDeleteAllConds_Click(object sender, System.EventArgs e)
        {
            this.btnFilterFinish.Enabled = true;
            this.btnFilterNext.Enabled = true;
            this.txtFilterByFile.Enabled = false;
            this.btnFilterByFileBrowse.Enabled = false;
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            ((frmDialog) this.ParentForm).Close();
        }

        private void btnFilterFinish_Click(object sender, System.EventArgs e)
        {
            this.m_strStateCountyPlotSQL = "";
            this.m_strStateCountySQL = "";
            this.m_intError = 0;
            InitializeDatasource();
            if (m_intError == 0)
            {
                //No specific conditions to remove. Delete all of them.
                if (this.rdoDeleteAllConds.Checked)
                {
                    /*TODO: either set m_strCondCNs to all possible cond.cn values 
                     * or modify the SQL to skip the where filters 
                     * (caution: this approach deletes tables that don't have 
                     * plot/cond/tree/stand identifiers at all)
                     */
                    DeleteCondsFromBiosumProject_Start();
                }
                else if (this.rdoFilterByMenu.Checked)
                {
                    //TODO: set m_strCondCNs via the GUI
                    DeleteCondsFromBiosumProject_Start();
                }
                else if (this.rdoFilterByFile.Checked)
                {
                    if (System.IO.File.Exists(this.txtFilterByFile.Text.Trim()) == true)
                    {
                        this.m_strCondCNs =
                            this.CreateDelimitedStringList(this.txtFilterByFile.Text.Trim(), ",", ",", false);
                        if (this.m_intError == 0)
                        {
                            DeleteCondsFromBiosumProject_Start();
                        }
                    }
                    else
                    {
                        MessageBox.Show("!!" + this.txtFilterByFile.Text.Trim() + " could not be found!!",
                            "Delete Conditions", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                    }
                }
            }
        }

        //throw new NotImplementedException("Delete conds, with or without filters, using m_strCondCNs built with either a text file or GUI menu. use it to collect biosum_cond_ids, plot.cn, tree.cn");
        private void DeleteCondsFromBiosumProject_Start()
        {
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oEventStopThread.Reset();
            frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            this.m_strCurrentProcess = "DeleteConditions";
            this.StartTherm("2", "Delete Biosum Condition Data");
            frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(DeleteCondsFromBiosumProject_Process));
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.m_oThread.Start();
        }

        private void DeleteCondsFromBiosumProject_Process()
        {
            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
            this.m_intError = 0;
            try
            {
                this.m_ado = new ado_data_access();
                this.m_dao = new dao_data_access();

                //progress bar 1: single process
                this.SetThermValue(m_frmTherm.progressBar1, "Maximum", 100);
                this.SetThermValue(m_frmTherm.progressBar1, "Minimum", 0);
                this.SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                this.SetLabelValue(m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) m_frmTherm, "Visible", true);
                //progress bar 2: overall progress
                this.SetThermValue(m_frmTherm.progressBar2, "Maximum", 100);
                this.SetThermValue(m_frmTherm.progressBar2, "Minimum", 0);
                this.SetThermValue(m_frmTherm.progressBar2, "Value", 0);
                this.SetLabelValue(m_frmTherm.lblMsg2, "Text", "Overall Progress");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) m_frmTherm, "Visible", true);


                this.m_strTempMDBFile = m_oDatasource.CreateMDBAndTableDataSourceLinks();
                //get an ado connection string for the temp mdb file
                this.m_strTempMDBFileConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile, "", "");
                //create a new connection to the temp MDB file
                this.m_connTempMDBFile = new System.Data.OleDb.OleDbConnection();
                //open the connection to the temp mdb file 
                this.m_ado.OpenConnection(this.m_strTempMDBFileConn, ref this.m_connTempMDBFile);

                UpdateProgressBar2(10);
                SetConditionLookupAndIDStrings();

                //Need to access plot.fvs_variants before deleting records
                string[] strVariants = m_oDatasource.getVariants();

                //ProjectRoot\db Section
                UpdateProgressBar2(20);
                ConnectToDatabasesInPathAndExecuteDeletes(Directory.GetFiles(m_strProjDir + "\\db\\", "*.mdb"), strTableExceptions: new string[] {"tree_regional_biomass"});

                //ProjectRoot\gis Section
                UpdateProgressBar2(30);
                ConnectToDatabasesInPathAndExecuteDeletes(Directory.GetFiles(m_strProjDir + "\\gis\\db\\", "*.mdb"));

                //ProjectRoot\processor Section
                UpdateProgressBar2(40);
                foreach (string subfolder in Directory.GetDirectories(m_strProjDir + "\\processor\\"))
                {
                    if (Path.GetFileName(subfolder) == "db")
                        ConnectToDatabasesInPathAndExecuteDeletes(Directory.GetFiles(subfolder, "*.mdb"));
                    else 
                        if (Directory.Exists(subfolder + "\\db\\"))
                            ConnectToDatabasesInPathAndExecuteDeletes(Directory.GetFiles(subfolder + "\\db\\", "*.mdb"));
                }

                //FVS Section
                UpdateProgressBar2(50);
                if (strVariants != null && strVariants.Length > 0)
                {
                    string strFvsDataDir = m_strProjDir + "\\fvs\\data\\";
                    int num_variants_in_fvs_data = 0;
                    foreach (string variant in strVariants)
                        if (Directory.Exists(strFvsDataDir + "\\" + variant + "\\"))
                            num_variants_in_fvs_data++;
                    int progressbar2_value = 50;
                    foreach (string variant in strVariants)
                    {
                        //Collect pathfiles of databases to delete from in FVS Data directory
                        string strVariantPath = strFvsDataDir + variant + "\\";
                        if (Directory.Exists(strVariantPath))
                        {
                            progressbar2_value += (30 / (num_variants_in_fvs_data + 1));
                            UpdateProgressBar2(progressbar2_value);
                            ConnectToDatabasesInPathAndExecuteDeletes(
                                Directory.GetFiles(strVariantPath, "*.*", SearchOption.AllDirectories)
                                    .Where(s => s.ToLower().EndsWith(".mdb") || s.ToLower().EndsWith(".accdb"))
                                    .ToArray(),
                                strTableExceptions: new string[] {"FVS_GroupAddFilesAndKeywords"});
                        }
                    }
                }

                //PREPOST Databases
                UpdateProgressBar2(80);
                ConnectToDatabasesInPathAndExecuteDeletes(
                    Directory.GetFiles(m_strProjDir + "\\fvs\\db\\", "PREPOST*.accdb"),
                    strTableExceptions: new string[] {"FVS_TREE"});
                UpdateProgressBar2(100);

                //Results
                if (Checked(chkCreateLog))
                {
                    CreateLogFile(frmMain.g_oEnv.strTempDir + "\\biosum_deleted_records" +
                                  String.Format("{0:yyyyMMdd}", DateTime.Now) + ".txt");
                }

                MessageBox.Show(
                    String.Format("Successfully deleted data associated with {0} Conditions!",
                        ((HashSet<string>) m_dictIdentityColumnsToValues["biosum_cond_id"][0]).Count),
                    "Delete Conditions Results");

                //Cleanup section, assuming no exceptions were thrown
                this.m_connTempMDBFile.Close();
                while (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
                    System.Threading.Thread.Sleep(1000);
                if (m_ado != null)
                {
                    if (m_ado.m_DataSet != null)
                    {
                        this.m_ado.m_DataSet.Clear();
                        this.m_ado.m_DataSet.Dispose();
                    }
                    this.m_ado = null;
                }
                if (m_dao != null)
                {
                    this.m_dao.m_DaoWorkspace.Close();
                    this.m_dao.m_DaoWorkspace = null;
                    this.m_dao = null;
                }

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) ReferenceFormDialog, "Visible",
                    true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) ReferenceFormDialog, "Enabled",
                    true);

                DeleteCondsFromBiosumProject_Finish();
            }
            catch (System.Threading.ThreadInterruptedException err)
            {
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (System.Threading.ThreadAbortException err)
            {
                if (this.m_connTempMDBFile != null)
                {
                    if (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
                    {
                        m_ado.CloseConnection(m_connTempMDBFile);
                    }
                    m_connTempMDBFile = null;
                }
                if (m_ado != null)
                {
                    if (m_ado.m_DataSet != null)
                    {
                        this.m_ado.m_DataSet.Clear();
                        this.m_ado.m_DataSet.Dispose();
                    }
                    this.m_ado = null;
                }
                this.ThreadCleanUp();
                this.CleanupThread();
            }
            catch (Exception err)
            {
                MessageBox.Show("!!Error!! \n" +
                                "Module - uc_delete_conditions:DeleteCondsFromBiosumProject_Process  \n" +
                                "Err Msg - " + err.Message.ToString().Trim(),
                    "FVS Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }
            finally
            {
            }

            if (this.m_connTempMDBFile != null)
            {
                if (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
                {
                    m_ado.CloseConnection(m_connTempMDBFile);
                }
                m_connTempMDBFile = null;
            }
            if (m_ado != null)
            {
                if (m_ado.m_DataSet != null)
                {
                    this.m_ado.m_DataSet.Clear();
                    this.m_ado.m_DataSet.Dispose();
                }
                this.m_ado = null;
            }
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) ReferenceFormDialog, "Enabled",
                true);
            if (this.m_frmTherm != null)
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) m_frmTherm, "Visible", false);


            DeleteCondsFromBiosumProject_Finish();

            CleanupThread();

            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
        }


        private void UpdateProgressBar1(string label, int value)
        {
            this.SetLabelValue(m_frmTherm.lblMsg, "Text", label);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control) this.m_frmTherm.progressBar1,
                "Value", value);
        }

        private void UpdateProgressBar2(int value)
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control) this.m_frmTherm.progressBar2,
                "Value", value);
        }

        private void ConnectToDatabasesInPathAndExecuteDeletes(string[] strDatabaseNames, string[] strTargetTables = null, string[] strTableExceptions = null)
        {
            int progressbar1_value_incr = 100 / (strDatabaseNames.Length + 1);
            int counter = 0;

            foreach (string db in strDatabaseNames)
            {
                UpdateProgressBar1("Deleting from " + Path.GetFileName(db), counter);
                counter += progressbar1_value_incr;
                ExecuteDeleteOnTables(db, tables: strTargetTables, exceptions: strTableExceptions);
            }
            UpdateProgressBar1("", 0);
        }

        private void SetConditionLookupAndIDStrings()
        {
            setBiosumCondCNs = new HashSet<string>();
            setBiosumCondIds = new HashSet<string>();
            setBiosumPlotIds = new HashSet<string>();
            setPlotCNs = new HashSet<string>();
            setTreeCNs = new HashSet<string>();

            if (this.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control) m_frmTherm, "AbortProcess"))
            {
                SetThermValue(m_frmTherm.progressBar1, "Value", 20);
                this.m_ado.CreateDataSet(this.m_connTempMDBFile,
                    String.Format(
                        "SELECT c.cn, c.biosum_cond_id, t.cn FROM ({0} c INNER JOIN {1} p ON c.biosum_plot_id=p.biosum_plot_id) LEFT JOIN {2} t on c.biosum_cond_id=t.biosum_cond_id WHERE c.cn IN ({3});",
                        m_strCondTable, m_strPlotTable, m_strTreeTable, m_strCondCNs), "identifiers");
                foreach (DataRow row in m_ado.m_DataSet.Tables["identifiers"].Rows)
                {
                    setBiosumCondCNs.Add(String.Format("'{0}'", row[0]));
                    setBiosumCondIds.Add(String.Format("'{0}'", row[1]));
                    setTreeCNs.Add(String.Format("'{0}'", row[2]));
                }
                m_ado.CreateDataSet(this.m_connTempMDBFile,
                    String.Format(
                        "SELECT allconds.biosum_plot_id, allconds.cn " + //, allconds.cntConds, someconds.cntConds " +
                        "FROM (SELECT p.biosum_plot_id, p.cn, count(*) as cntConds FROM {0} c INNER JOIN {1} p ON c.biosum_plot_id = p.biosum_plot_id WHERE c.cn IN ({2}) GROUP BY p.biosum_plot_id, p.cn) someconds " +
                        "RIGHT JOIN (SELECT p.biosum_plot_id, p.cn, count(*) as cntConds FROM {0} c INNER JOIN {1} p ON c.biosum_plot_id = p.biosum_plot_id GROUP BY p.biosum_plot_id, p.cn) allconds " +
                        "ON allconds.biosum_plot_id=someconds.biosum_plot_id WHERE allconds.cntConds=someconds.cntConds",
                        m_strCondTable, m_strPlotTable, m_strCondCNs), "plots_with_all_conds_deleted");
                foreach (DataRow row in m_ado.m_DataSet.Tables["plots_with_all_conds_deleted"].Rows)
                {
                    setBiosumPlotIds.Add(String.Format("'{0}'", row[0]));
                    setPlotCNs.Add(String.Format("'{0}'", row[1]));
                }
                m_intError = m_ado.m_intError;
            }

            //Create strings from the HashSets to insert into SQL filters 
            m_strBiosumCondCNs = CreateCommaDelimitedString(setBiosumCondCNs);
            m_strBiosumCondIds = CreateCommaDelimitedString(setBiosumCondIds);
            m_strBiosumPlotIds = CreateCommaDelimitedString(setBiosumPlotIds);
            m_strPlotCNs = CreateCommaDelimitedString(setPlotCNs);
            if (setTreeCNs.Count < 5000)
            {
                m_strTreeCNs = CreateCommaDelimitedString(setTreeCNs);
            }

            m_dictIdentityColumnsToValues = new Dictionary<string, object[]>
            {
                {"biosum_cond_id", new object[] {setBiosumCondCNs, m_strBiosumCondIds}},
                {"StandID", new object[] {setBiosumCondCNs, m_strBiosumCondIds}},
                {"Stand_ID", new object[] {setBiosumCondCNs, m_strBiosumCondIds}},
                {"biosum_plot_id", new object[] {setBiosumPlotIds, m_strBiosumPlotIds}},
                //Only usage of cnd_cn is in FCS schema, which uses biosum_cond_id values
                {"cnd_cn", new object[] {setBiosumCondCNs, m_strBiosumCondCNs}},
                //plt_cn used in two tables. pop_plot_stratum_assgn uses plot.cn, fcs uses biosum_plot_id
                //{"plt_cn", new object[] {setPlotCNs, m_strPlotCNs}},
                //used in master.tree_regional_biomass. values are tree.cn
                //{"tre_cn", new object[] {setTreeCNs, m_strTreeCNs}},
            };
        }


        private void ExecuteDeleteOnTables(string strDbPathFile, string[] tables = null, string[] exceptions = null)
        {
            using (var conn = new OleDbConnection(m_ado.getMDBConnString(strDbPathFile, "", "")))
            {
                conn.Open();

                string[] strTables = tables;
                if (tables == null || tables.Length == 0)
                {
                    strTables = m_ado.getTableNamesOfSpecificTypes(conn);
                    //In case none of the tables are valid
                    if (strTables.Length == 1 && strTables[0] == "")
                        return;
                }

                string column;
                foreach (string table in strTables)
                {
                    if (exceptions != null && exceptions.Contains(table))
                    {
                        continue;
                    }

                    column = null;
                    foreach (string col in m_dictIdentityColumnsToValues.Keys)
                    {
                        if (m_ado.ColumnExist(conn, table, col))
                        {
                            column = col;
                            break;
                        }
                    }

                    if (!(String.IsNullOrEmpty(column)) &&
                        ((HashSet<string>) m_dictIdentityColumnsToValues[column][0]).Count > 0)
                    {
                        m_ado.AddIndex(conn, table, column + "_delete_idx", column);
                        int deletedRecords = BuildAndExecuteDeleteSQLStmts(conn, table, column);
                        AddDeletedCountToDictionary(strDbPathFile, table, deletedRecords);
                        m_ado.SqlNonQuery(conn, String.Format("DROP INDEX {0} ON {1}", column + "_delete_idx", table));
                    }
                }
            }
            m_dao.CompactMDB(strDbPathFile);
        }

        private void AddDeletedCountToDictionary(string strDbPathFile, string table, int deletedRecords)
        {
            if (!m_dictDeletedRowCountsByDatabaseAndTable.Keys.Contains(strDbPathFile))
            {
                m_dictDeletedRowCountsByDatabaseAndTable.Add(strDbPathFile,
                    new Dictionary<string, int>());
            }
            m_dictDeletedRowCountsByDatabaseAndTable[strDbPathFile].Add(table, deletedRecords);
        }

        private int BuildAndExecuteDeleteSQLStmts(OleDbConnection oConn, string table, string column)
        {
            int deletedRecords = 0;
            string column_key = column;
            string strSQL;
            string[] edgeCaseTables = new string[] {"fcs_biosum_volumes_input"};

            if (edgeCaseTables.Contains(table))
            {
                //FCS schema uses cnd_cn/plt_cn, but we set these values to biosum_cond_id/biosum_plot_id in FVSOUT stage
                if (table.ToLower() == "fcs_biosum_volumes_input" && column.ToLower() == "cnd_cn")
                    column_key = "biosum_cond_id";
            }

            if (((HashSet<string>) m_dictIdentityColumnsToValues[column_key][0]).Count < 5000)
            {
                deletedRecords += (int) m_ado.getSingleDoubleValueFromSQLQuery(oConn, String.Format(
                    "SELECT COUNT(*) FROM {0} WHERE {1} IN ({2});", table, column,
                    m_dictIdentityColumnsToValues[column_key][1]), table);
                strSQL = String.Format("DELETE FROM {0} WHERE {1} IN ({2});", table, column,
                    m_dictIdentityColumnsToValues[column_key][1]);
                m_ado.SqlNonQuery(oConn, strSQL);
            }

            else
            {
                foreach (string id in (HashSet<string>) m_dictIdentityColumnsToValues[column_key][0])
                {
                    deletedRecords += (int) m_ado.getSingleDoubleValueFromSQLQuery(oConn,
                        String.Format("SELECT COUNT(*) FROM {0} WHERE {1} = {2};", table, column, id), table);
                    strSQL = String.Format("DELETE FROM {0} WHERE {1} = {2};", table, column, id);
                    m_ado.SqlNonQuery(oConn, strSQL);
                }
            }

            return deletedRecords;
        }

        private void CreateLogFile(string strPathFile)
        {
            int deletedRecordsCount;
            foreach (var strDbPathFile in m_dictDeletedRowCountsByDatabaseAndTable.Keys)
            {
                frmMain.g_oUtils.WriteText(strPathFile, String.Concat(strDbPathFile, Environment.NewLine));
                foreach (var strTable in m_dictDeletedRowCountsByDatabaseAndTable[strDbPathFile].Keys)
                {
                    deletedRecordsCount = m_dictDeletedRowCountsByDatabaseAndTable[strDbPathFile][strTable];
                    if (deletedRecordsCount > 0)
                    {
                        frmMain.g_oUtils.WriteText(strPathFile,
                            String.Format("  {0} records deleted from {1}{2}", deletedRecordsCount, strTable,
                                Environment.NewLine));
                    }
                }
            }
        }

        private string CreateCommaDelimitedString(ICollection<string> strs)
        {
            StringBuilder stringBuilder = new StringBuilder();
            bool bFirst = true;
            foreach (string str in strs)
            {
                if (bFirst)
                {
                    stringBuilder.Append(str);
                    bFirst = false;
                }
                else
                {
                    stringBuilder.Append("," + str);
                }
            }
            return stringBuilder.ToString();
        }

        private void CleanupThread()
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) ReferenceFormDialog, "Visible",
                true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) ReferenceFormDialog, "Enabled",
                true);
        }

        private void ThreadCleanUp()
        {
            try
            {
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) ReferenceFormDialog, "Visible",
                    true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) ReferenceFormDialog, "Enabled",
                    true);

                if (this.m_frmTherm != null)
                {
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form) m_frmTherm, "Close");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form) m_frmTherm, "Dispose");

                    this.m_frmTherm = null;
                }
            }
            catch
            {
            }
        }

        private void DeleteCondsFromBiosumProject_Finish()
        {
            if (this.m_frmTherm != null)
            {
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Close");
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Dispose");
                this.m_frmTherm = null;
            }
            this.m_strCurrentProcess = "";
            frmMain.g_oDelegate.SetControlPropertyValue(this, "Enabled", true);
            ((frmDialog) ParentForm).MinimizeMainForm = false;
        }


        private void StartTherm(string p_strNumberOfTherms, string p_strTitle)
        {
            this.m_frmTherm = new frmTherm((frmDialog) this.ParentForm, p_strTitle);

            this.m_frmTherm.Text = p_strTitle;
            this.m_frmTherm.lblMsg.Text = "";
            this.m_frmTherm.lblMsg2.Text = "";
            this.m_frmTherm.Visible = false;
            this.m_frmTherm.btnCancel.Visible = false; //disallow users to cancel delete process
            this.m_frmTherm.btnCancel.Enabled = false;
            this.m_frmTherm.lblMsg.Visible = true;
            this.m_frmTherm.progressBar1.Minimum = 0;
            this.m_frmTherm.progressBar1.Visible = true;
            this.m_frmTherm.progressBar1.Maximum = 10;

            if (p_strNumberOfTherms == "2")
            {
                this.m_frmTherm.progressBar2.Size = this.m_frmTherm.progressBar1.Size;
                this.m_frmTherm.progressBar2.Left = this.m_frmTherm.progressBar1.Left;
                this.m_frmTherm.progressBar2.Top =
                    Convert.ToInt32(this.m_frmTherm.progressBar1.Top + (this.m_frmTherm.progressBar1.Height * 3));
                this.m_frmTherm.lblMsg2.Top =
                    this.m_frmTherm.progressBar2.Top + this.m_frmTherm.progressBar2.Height + 5;
                this.m_frmTherm.Height = this.m_frmTherm.lblMsg2.Top + this.m_frmTherm.lblMsg2.Height +
                                         this.m_frmTherm.btnCancel.Height + 50;
                this.m_frmTherm.btnCancel.Top =
                    this.m_frmTherm.ClientSize.Height - this.m_frmTherm.btnCancel.Height - 5;
                this.m_frmTherm.lblMsg2.Show();
                this.m_frmTherm.progressBar2.Visible = true;
            }
            this.m_frmTherm.AbortProcess = false;
            this.m_frmTherm.Refresh();
            this.m_frmTherm.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            ((frmDialog)this.ParentForm).Enabled=false;
            this.m_frmTherm.Visible = true;
        }

        public void StopThread()
        {
            string strMsg = "";

            frmMain.g_oDelegate.AbortProcessing("FIA Biosum", "Do you wish to cancel adding plot data?");

            if (frmMain.g_oDelegate.CurrentThreadProcessAborted)
            {
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) m_frmTherm, "AbortProcess",
                    true);
//                this.CancelThreadCleanup();
//                this.ThreadCleanUp();
                throw new NotImplementedException();
            }
        }

        private bool Checked(System.Windows.Forms.RadioButton p_rdoButton)
        {
            return (bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.RadioButton)p_rdoButton, "Checked", false);
        }
        private bool Checked(System.Windows.Forms.CheckBox p_chkBox)
        {
            return (bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.CheckBox)p_chkBox, "Checked", false);
        }

        private void SetThermValue(System.Windows.Forms.ProgressBar p_oPb, string p_strPropertyName, int p_intValue)
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control) p_oPb, p_strPropertyName,
                (int) p_intValue);
        }

        private int GetThermValue(System.Windows.Forms.ProgressBar p_oPb, string p_strPropertyName)
        {
            return (int) frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control) p_oPb,
                p_strPropertyName, false);
        }

        private bool GetBooleanValue(System.Windows.Forms.Control p_oControl, string p_strPropertyName)
        {
            return (bool) frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control) p_oControl,
                p_strPropertyName, false);
        }


        private void SetLabelValue(System.Windows.Forms.Label p_oLabel, string p_strPropertyName, string p_strValue)
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Label) p_oLabel, p_strPropertyName,
                p_strValue);
        }

        public frmDialog ReferenceFormDialog { set; get; }


        private void btnFilterByFileBrowse_click(object sender, EventArgs e)
        {
            var OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Title = "Text File With PLOT_CN data";
            OpenFileDialog1.Filter = "Text File (*.TXT;*.DAT) |*.txt;*.dat";
            var result = OpenFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                if (OpenFileDialog1.FileName.Trim().Length > 0)
                    txtFilterByFile.Text = OpenFileDialog1.FileName.Trim();
                OpenFileDialog1 = null;
            }
        }

        /// <summary>
        /// create a delimited string list from a text file
        /// that has a single column of data with multiple rows
        /// </summary>
        /// <param name="p_strTxtFile">text file containing the column of data</param>
        /// <param name="p_strTxtFileDelimiter">specified character between list items</param>
        /// <param name="p_strListDelimiter">specified character between list items</param>
        /// <param name="p_bNumericDataType">specifies if the column data to retrieve in the text file is numeric</param>
        /// <returns></returns>
        private string CreateDelimitedStringList(string p_strTxtFile, string p_strTxtFileDelimiter,
            string p_strListDelimiter, bool p_bNumericDataType)
        {
            //The DataSet to Return
            //DataSet result = new DataSet();
            this.m_intError = 0;
            string strList = "";
            string str = "";
            try
            {
                //Open the file in a stream reader.
                System.IO.StreamReader s = new System.IO.StreamReader(p_strTxtFile);
                //Read the rest of the data in the file.        
                string AllData = s.ReadToEnd();

                //Split off each row at the Carriage Return/Line Feed
                //Default line ending in most <A class=iAs style="FONT-WEIGHT: normal; FONT-SIZE: 100%; PADDING-BOTTOM: 1px; COLOR: darkgreen; BORDER-BOTTOM: darkgreen 0.07em solid; BACKGROUND-COLOR: transparent; TEXT-DECORATION: underline" href="#" target=_blank itxtdid="2592535">windows</A> exports.  
                string[] rows = AllData.Split("\r\n".ToCharArray());

                //Now add each row to the DataSet        
                foreach (string r in rows)
                {
                    //Split the row at the delimiter.
                    string[] items = r.Split(p_strTxtFileDelimiter.ToCharArray());
                    str = items[0].Trim(); //plot_cn in first column
                    str = str.Replace("\"", ""); //remove any quotations
                    if (str.Trim().Length > 0)
                    {
                        if (strList.Trim().Length == 0)
                        {
                            if (p_bNumericDataType == true)
                            {
                                strList = str.Trim();
                            }
                            else
                            {
                                strList = "'" + str.Trim() + "'";
                            }
                        }
                        else
                        {
                            if (p_bNumericDataType == true)
                            {
                                strList = strList + p_strListDelimiter.Trim() + str.Trim();
                            }
                            else
                            {
                                strList = strList + p_strListDelimiter.Trim() + "'" + str.Trim() + "'";
                            }
                        }
                    }
                }
            }
            catch (Exception caught)
            {
                this.m_intError = -1;
                MessageBox.Show("!!Error: CreateDelimitedStringList() Routine Error Msg:" + caught.Message);
            }
            return strList;
        }

        private void txtFilterByFile_TextChanged(object sender, EventArgs e)
        {
            this.btnFilterFinish.Enabled = true;
        }
    }
}
