using System;
using System.Data;
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
		private string m_strBiosumCondIds="";

        //TODO: Help files
        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultDatabaseXPSFile;
        private int m_intError;
        private string m_strStateCountyPlotSQL;
        private string m_strStateCountySQL;

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
            var strProjDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();

            m_oDatasource = new Datasource();
            m_oDatasource.LoadTableColumnNamesAndDataTypes = false;
            m_oDatasource.LoadTableRecordCount = false;
            m_oDatasource.m_strDataSourceMDBFile = strProjDir.Trim() + "\\db\\project.mdb";
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
            ((frmDialog)this.ParentForm).Close();
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
                    MessageBox.Show("You really want to delete everything here?");
//                    throw new NotImplementedException("This should eventually delete ALL conds/plots/trees through the FVSOUT phase of BioSum.");
                    //delete pretty much everything. building a string of 1+ biosum_cond_ids is skipped (filter by file radio button not selected)
                    DeleteCondsFromBiosumProject();
                }
                else if (this.rdoFilterByMenu.Checked)
                {
                    MessageBox.Show("Deleting conds based on menu selection...");
//                    throw new NotImplementedException("Eventually deletes specific conds/plots/trees through the FVSOUT phase of BioSum.");
                    DeleteCondsFromBiosumProject();
                }
                else if (this.rdoFilterByFile.Checked)
                {
                    if (System.IO.File.Exists(this.txtFilterByFile.Text.Trim()) == true)
                    {
                        this.m_strBiosumCondIds = this.CreateDelimitedStringList(this.txtFilterByFile.Text.Trim(), ",", ",", false);
                        if (this.m_intError == 0)
                        {
                            //this.LoadMDBPlotCondTreeData_Start();
                            MessageBox.Show("This would be deleting information related to the following conditions:" + m_strBiosumCondIds);
//                            throw new NotImplementedException("This should eventually delete SPECIFIC conds/plots/trees through the FVSOUT phase of BioSum.");
                            DeleteCondsFromBiosumProject();
                        }
                    }
                    else
                    {
                        MessageBox.Show("!!" + this.txtFilterByFile.Text.Trim() + " could not be found!!", "Delete Conditions", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    }
                }
            }
        }

        private void DeleteCondsFromBiosumProject()
        {
            //throw new NotImplementedException("Delete conds, with or without filters, using m_strBiosumCondIds built with either a text file or GUI menu. use it to collect biosum_cond_ids, plot.cn, tree.cn");
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
		private string CreateDelimitedStringList(string p_strTxtFile,string p_strTxtFileDelimiter, string p_strListDelimiter,bool p_bNumericDataType)
		{
			//The DataSet to Return
			//DataSet result = new DataSet();
			this.m_intError=0;
			string strList="";
			string str="";
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
				foreach(string r in rows)
				{
					//Split the row at the delimiter.
					string[] items = r.Split(p_strTxtFileDelimiter.ToCharArray());
					str = items[0].Trim();  //plot_cn in first column
					str = str.Replace("\"",""); //remove any quotations
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
				this.m_intError=-1;
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