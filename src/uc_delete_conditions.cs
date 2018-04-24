using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

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
        //TODO: more strings to hold FVSOUT table information?

        //TODO:for collecting biosum_cond_id values from text file
        private string m_strCondCNs;
        private string m_strBiosumCondCNs;
        private string m_strBiosumCondIds;
        private string m_strBiosumPlotIds;
        private string m_strPlotCNs;
        private string m_strTreeCNs;
        Dictionary<string, string> m_dictIDColumnsToValues;

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
        private OleDbConnection m_connTempMDBFile;
        private string m_strTempMDBFileConn;
        private string m_strTempMDBFile;
        private string m_strSQL;
        private string m_strProjDir;

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

            m_oEnv = new env();
        }

        private void InitializeDatasource()
        {
            m_strProjDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();

            m_oDatasource = new Datasource();
            m_oDatasource.LoadTableColumnNamesAndDataTypes = false;
            m_oDatasource.LoadTableRecordCount = false;
            m_oDatasource.m_strDataSourceMDBFile = m_strProjDir.Trim() + "\\db\\project.mdb";
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

            //TODO: Delete Conditions, parent plots if only one cond, the cond's trees, etc. all the way through FVSOUT
            if (m_intError == 0)
            {
                //No specific conditions to remove. Delete all of them.
                if (this.rdoDeleteAllConds.Checked)
                {
                    //LoadMDBPlotCondTreeData_Start();
                    //MessageBox.Show("You really want to delete everything here?");
                    //throw new NotImplementedException("This should eventually delete ALL conds/plots/trees through the FVSOUT phase of BioSum.");
                    //delete pretty much everything. building a string of 1+ biosum_cond_ids is skipped (filter by file radio button not selected)
                    DeleteCondsFromBiosumProject_Start();
                }
                else if (this.rdoFilterByMenu.Checked)
                {
                    //MessageBox.Show("Deleting conds based on menu selection...");
                    //throw new NotImplementedException("Eventually deletes specific conds/plots/trees through the FVSOUT phase of BioSum.");
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
                            //this.LoadMDBPlotCondTreeData_Start();
                            //MessageBox.Show("This would be deleting information related to the following conditions:" + m_strCondCNs);
                            //throw new NotImplementedException("This should eventually delete SPECIFIC conds/plots/trees through the FVSOUT phase of BioSum.");
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

            //Close the form after deleting condititions finished
            //((frmDialog) this.ParentForm).Close(); //causes premature disposal with multithreaded applications
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

            //-----------PREPARATION FOR DELETING COND RECORDS---------//

            try
            {
                this.m_ado = new ado_data_access();

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
                this.SetLabelValue(m_frmTherm.lblMsg, "Text", "Collecting information about conditions to delete..."); //todo: rename step label
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control) this.m_frmTherm.progressBar1,
                    "Value", 10);

                this.m_strTempMDBFile = m_oDatasource.CreateMDBAndTableDataSourceLinks();
                //get an ado connection string for the temp mdb file
                this.m_strTempMDBFileConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile, "", "");
                //create a new connection to the temp MDB file
                this.m_connTempMDBFile = new System.Data.OleDb.OleDbConnection();
                //open the connection to the temp mdb file 
                this.m_ado.OpenConnection(this.m_strTempMDBFileConn, ref this.m_connTempMDBFile);


                SetConditionLookupAndIDStrings();

                //FVS Section
                string[] strVariants = m_oDatasource.getVariants();
                string strFvsDataDir = m_strProjDir + "\\fvs\\data\\";
                foreach (string variant in strVariants)
                {
                    //Collect pathfiles of databases to delete from in FVS Data directory
                    string strVariantPath = strFvsDataDir + variant + "\\";

                    //TODO: Check if this variant folder exists
                    if (Directory.Exists(strVariantPath))
                    {
                        //FVSIN
                        string strFvsIn = strVariantPath + "fvsin.accdb";
                        strFvsIn = File.Exists(strFvsIn) ? strFvsIn : null;

                        //FVSOUT
                        string[] strFvsOutMDBs = Directory.GetFiles(strVariantPath, "fvsout*.mdb");
                        string[] strFvsOutBiosumACCDBs = Directory.GetFiles(strVariantPath, "fvsout*biosum.accdb");
                        string[] strBiosumCalcMDBs = Directory.GetFiles(strVariantPath + "BiosumCalc\\", "*TREE_CUTLIST.mdb");


                        //Connect to each database

                        //ExecuteDeleteOnAllTables(db);
                        //tables and exceptions will be mutually exclusive because they're not compared to the tables inside the database if tables != null
                        ExecuteDeleteOnTables(strFvsIn, exceptions: new string[] {"FVS_GroupAddFilesAndKeywords"});
//                        ExecuteDeleteOnTables(strFvsIn, tables: new string[] {"FVS_StandInit", "FVS_TreeInit", "FVS_PlotInit"});

                        foreach (string db in strFvsOutMDBs)
                        {
                            //ExecuteDeleteOnAllTables(db);
                            ExecuteDeleteOnTables(db);
                        }
                        foreach (string db in strFvsOutBiosumACCDBs)
                        {
                            //ExecuteDeleteOnAllTables(db);
                            ExecuteDeleteOnTables(db);
                            //fvsout*biosum.accdb has links to tables that might not exist in output table...
                        }
                        foreach (string db in strBiosumCalcMDBs)
                        {
                            //ExecuteDeleteOnAllTables(db);
                            ExecuteDeleteOnTables(db);
                        }
                    }
                }

                //PREPOST Databases
                string strFvsDbDir = m_strProjDir + "\\fvs\\db\\";
                string[] strPrePostDbs = Directory.GetFiles(strFvsDbDir, "PREPOST*.accdb");
                foreach (string db in strPrePostDbs)
                {
                    //One of the linked tables points to a temporary database. If you run the FVSOUT module, this code works because the linked tables aren't broken.
                    //TODO: Pass specific table names to this function so you don't open the ones that link to nonexistent databases.
                    //ExecuteDeleteOnAllTables(db);
                    ExecuteDeleteOnTables(db, exceptions: new string[] {"FVS_TREE"});
                }


//                //TODO: make this a function to call repeatedly with different m_strSQL values, connections, and progress bar values
//                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control) m_frmTherm, "AbortProcess"))
//                {
//                    SetThermValue(m_frmTherm.progressBar1, "Value", 80);
//                    m_ado.m_strSQL = "SELECT 1;"; //todo: delete from table where ___ in () function
//                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
//                    m_intError = m_ado.m_intError;
//                }



//                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control) m_frmTherm, "AbortProcess"))
//                {
//                }
//                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control) m_frmTherm, "AbortProcess"))
//                {
//                }
//                else
//                {
//                    MessageBox.Show("Some error occured in deleting conditions.");
//                }


                this.m_connTempMDBFile.Close();
                while (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
                    System.Threading.Thread.Sleep(1000);
                //this.m_ado.m_DataSet.Clear(); //todo: if m_DataSet is null, clear has null reference exception thrown
                //this.m_ado.m_DataSet.Dispose();
                this.m_ado = null;

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
//                this.CancelThreadCleanup();
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

        private void SetConditionLookupAndIDStrings()
        {
            System.Data.DataTable dtIdentifiers = new DataTable();
            HashSet<string> setBiosumCondCNs = new HashSet<string>();
            HashSet<string> setBiosumCondIds = new HashSet<string>();
            HashSet<string> setBiosumPlotIds = new HashSet<string>();
            HashSet<string> setPlotCNs = new HashSet<string>();
            HashSet<string> setTreeCNs = new HashSet<string>();


            if (this.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control) m_frmTherm, "AbortProcess"))
            {
                SetThermValue(m_frmTherm.progressBar1, "Value", 20);
                this.m_ado.CreateDataSet(this.m_connTempMDBFile,
                    String.Format(
                        "SELECT c.cn, c.biosum_cond_id, c.biosum_plot_id, p.cn, t.cn FROM ({0} c INNER JOIN {1} p ON c.biosum_plot_id=p.biosum_plot_id) INNER JOIN {2} t on c.biosum_cond_id=t.biosum_cond_id WHERE c.cn IN ({3});",
                        m_strCondTable, m_strPlotTable, m_strTreeTable, m_strCondCNs), "identifiers");
                dtIdentifiers = this.m_ado.m_DataSet.Tables["identifiers"];
                foreach (DataRow row in dtIdentifiers.Rows)
                {
                    setBiosumCondCNs.Add(String.Format("'{0}'", row[0]));
                    setBiosumCondIds.Add(String.Format("'{0}'", row[1]));
                    setBiosumPlotIds.Add(String.Format("'{0}'", row[2]));
                    setPlotCNs.Add(String.Format("'{0}'", row[3]));
                    setTreeCNs.Add(String.Format("'{0}'", row[4]));
                }
                m_intError = m_ado.m_intError;
            }

            //Create strings from the HashSets to insert into SQL filters 
            m_strBiosumCondCNs = CreateCommaDelimitedString(setBiosumCondCNs);
            m_strBiosumCondIds = CreateCommaDelimitedString(setBiosumCondIds);
            m_strBiosumPlotIds = CreateCommaDelimitedString(setBiosumPlotIds);
            m_strPlotCNs = CreateCommaDelimitedString(setPlotCNs);
            m_strTreeCNs = CreateCommaDelimitedString(setTreeCNs);

            m_dictIDColumnsToValues = new Dictionary<string, string>
            {
                {"biosum_cond_id", m_strBiosumCondIds},
                {"StandID", m_strBiosumCondIds},
                {"Stand_ID", m_strBiosumCondIds},
                {"biosum_plot_id", m_strBiosumPlotIds},
                {"plt_cn", m_strPlotCNs},
                {"cnd_cn", m_strBiosumCondCNs},
                {"tre_cn", m_strTreeCNs},
            };
        }

        private void ExecuteDeleteOnTables(string strDbPathFile, string[] tables = null, string[] exceptions = null)
        {
            //TODO: Logic for when rdoDeleteAllConds is selected
            OleDbConnection oConn = new OleDbConnection();
            m_ado.OpenConnection(m_ado.getMDBConnString(strDbPathFile, "", ""), ref oConn);

            string[] strTables = tables;
            if (tables == null || tables.Length == 0)
            {
                strTables = m_ado.getTableNames(oConn);
            }

            string strSQL, column;
            foreach (string table in strTables)
            {
                //in case a list of tables should be skipped
                if (exceptions != null && exceptions.Contains(table))
                {
                    using (StreamWriter file = new StreamWriter(m_strProjDir +"\\executedDeleteCondSQL.txt", true)) { 
                        file.WriteLine(String.Format("{0}\n\t{1}\n\tSKIPPED USING EXCEPTIONS ARRAY\n\n", strDbPathFile, table));
                    }
                    continue;
                }
                column = null;
                //Determine the column
                foreach (string col in m_dictIDColumnsToValues.Keys)
                {
                    if (m_ado.ColumnExist(oConn, table, col))
                    {
                        column = col;
                        break;
                    }
                }
                if (!String.IsNullOrEmpty(column))
                {
                    //Pick which column/identifier string to use in the strSQL
                    strSQL = String.Format("DELETE FROM {0} WHERE {1} IN ({2})", table, column, m_dictIDColumnsToValues[column]);
                    //TODO: delete conditions from this table 
                    using (StreamWriter file = new StreamWriter(m_strProjDir +"\\executedDeleteCondSQL.txt", true)) { 
                        file.WriteLine(String.Format("{0}\n\t{1}\n\t{2}\n\n", strDbPathFile, table, strSQL));
                    }
                }
            }
            m_ado.CloseConnection(oConn);
        }

//        private void ExecuteDeleteOnAllTables(string strDbPathFile)
//        {
//            //TODO: Logic for when rdoDeleteAllConds is selected
//            OleDbConnection oConn = new OleDbConnection();
//            m_ado.OpenConnection(m_ado.getMDBConnString(strDbPathFile, "", ""), ref oConn);
//            string[] tables;
//            tables = m_ado.getTableNames(oConn);
//            string strSQL, column;
//            foreach (string table in tables)
//            {
//                column = null;
//                //Determine the column
//                foreach (string col in m_dictIDColumnsToValues.Keys)
//                {
//                    if (m_ado.ColumnExist(oConn, table, col))
//                    {
//                        column = col;
//                        break;
//                    }
//                }
//                if (!String.IsNullOrEmpty(column))
//                {
//                    //Pick which column/identifier string to use in the strSQL
//                    strSQL = String.Format("DELETE FROM {0} WHERE {1} IN ({2})", table, column, m_dictIDColumnsToValues[column]);
//                    //TODO: delete conditions from this table 
//                    using (StreamWriter file = new StreamWriter(m_strProjDir +"\\executedDeleteCondSQL.txt", true)) { 
//                        file.WriteLine(String.Format("{0}\t{1}\t{2}\n", strDbPathFile, table, strSQL));
//                    }
//                }
//            }
//            m_ado.CloseConnection(oConn);
//        }

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
            // ((frmDialog)this.ParentForm).m_frmMain.Visible = true;
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) ReferenceFormDialog, "Visible",
                true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) ReferenceFormDialog, "Enabled",
                true);
        }

        private void ThreadCleanUp()
        {
            try
            {
                // ((frmDialog)this.ParentForm).m_frmMain.Visible = true;
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
//            this.m_strPlotIdList = "";


            if (this.m_frmTherm != null)
            {
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Close");
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Dispose");
                this.m_frmTherm = null;
            }
            if (m_intError != 0)
            {
//                this.m_strLoadedPopEstUnitInputTable = "";
//                this.m_strLoadedPopStratumInputTable = "";
//                this.m_strLoadedPpsaInputTable = "";
//                this.m_strLoadedFiadbInputFile = "";
            }
            else
            {
//                this.m_strLoadedPopEstUnitTxtInputFile = "";
//                this.m_strLoadedPopEvalTxtInputFile = "";
//                this.m_strLoadedPopStratumTxtInputFile = "";
//                this.m_strLoadedPpsaTxtInputFile = "";
            }
            this.m_strCurrentProcess = "";
            frmMain.g_oDelegate.SetControlPropertyValue(this, "Enabled", true);
            //((frmDialog) this.ParentForm).MinimizeMainForm = false;
        }


        private void StartTherm(string p_strNumberOfTherms, string p_strTitle)
        {
            this.m_frmTherm = new frmTherm((frmDialog) this.ParentForm, p_strTitle);

            this.m_frmTherm.Text = p_strTitle;
            this.m_frmTherm.lblMsg.Text = "";
            this.m_frmTherm.lblMsg2.Text = "";
            this.m_frmTherm.Visible = false;
            this.m_frmTherm.btnCancel.Visible = true;
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
            //((frmDialog)this.ParentForm).Enabled=false;
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
            return (bool) frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.RadioButton) p_rdoButton,
                "Checked", false);
        }

        private bool Checked(System.Windows.Forms.CheckBox p_chkBox)
        {
            return (bool) frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.CheckBox) p_chkBox,
                "Checked", false);
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