using System;
using System.IO;
using System.Windows.Forms;
using System.Collections;
using System.ComponentModel;
using System.Collections.Generic;


namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for fvs_input.
	/// </summary>
	public class fvs_input
	{ 
		public ado_data_access m_ado;
		public dao_data_access m_dao;
		public int m_intError=0;
		private string m_strPlotTable;
		private string m_strCondTable;
		private string m_strTreeTable;
		private string m_strTreeSpcTable;
		private string m_strSiteTreeTable;
		private string m_strFVSTreeSpcTable;
		
		private Datasource m_DataSource;
		private string m_strProjDir;
		private string m_strFVSInMDBFile;
		public string m_strTempMDBFile;
		private string m_strVariant;
		private frmTherm m_frmTherm;
		private string m_strInDir;
		private string m_strConn;
		private System.Data.DataTable m_dt;
		private string m_strFVSTreeIdIDBWorkTable;
		private string m_strFVSTreeIdFIAWorkTable;
        private string m_strDebugFile = frmMain.g_oEnv.strTempDir + "\\biosum_fvs_input_debug.txt";

        // Constants for site index equation table
        const String SI_DELIM = "|";
        const String SI_EMPTY = "@";

		/*******************************************************
		 **RECORD TYPE A
		 *******************************************************/
		private string[] m_strSlfALineArray;
		/// <summary>
		/// 0
		/// </summary>
		const int A_RECORDTYPE=0;
		/// <summary>
		/// 1
		/// </summary>
		const int A_STANDID=1;
		/// <summary>
		/// combination of state,plot,and condition id
		/// 2 
		/// </summary>
		const int A_FVSTREEFILENAME=2;
		/// <summary>
		/// 3
		/// </summary>
		const int A_SAMPLEPOINTSITEDATAFLAG=3;
		/// <summary>
		/// 4
		/// </summary>
		const int A_FVSVARIANTSUSED=4;
		/*******************************************************
		 **RECORD TYPE B
		 *******************************************************/
		private string[] m_strSlfBLineArray;
		/// <summary>
		/// 0
		/// </summary>
		const int B_RECORDTYPE=0;
		/// <summary>
		/// 1
		/// </summary>
		const int B_STANDID=1;
		/// <summary>
		/// 2
		/// </summary>
		const int B_INVYR=2;
		/// <summary>
		/// 3
		/// </summary>
		const int B_LAT=3;
		/// <summary>
		/// 4
		/// </summary>
		const int B_LON=4;
		/// <summary>
		/// 5
		/// </summary>
		const int B_FORESTCD=5;
		/// <summary>
		/// 6
		/// </summary>
		const int B_PLANTASSOC=6;
		/// <summary>
		/// 7
		/// </summary>
		const int B_STANDYEAROFORIGIN=7;
		/// <summary>
		/// 8
		/// </summary>
		const int B_ASPECT=8;
		/// <summary>
		/// 9
		/// </summary>
		const int B_SLOPEPERCENT=9;
		/// <summary>
		/// 10
		/// </summary>
		const int B_ELEV=10;
		/// <summary>
		/// 11
		/// </summary>
		const int B_BASALAREAFACTOR=11;
		/// <summary>
		/// 12
		/// </summary>
		const int B_INVERSESMALLTREE=12;
		/// <summary>
		/// 13
		/// </summary>
		const int B_DBHBREAKPOINT=13;
		/// <summary>
		/// 14
		/// </summary>
		const int B_NUMBEROFPLOTSPERFILE=14;
		/// <summary>
		/// 15
		/// </summary>
		const int B_NUMBEROFNONSTOCKABLEPLOTS=15;
		/// <summary>
		/// 16
		/// </summary>
		const int B_STANDSAMPLINGWEIGHT=16;
		/// <summary>
		/// 17
		/// </summary>
		const int B_STOCKABLEPERCENT=17;
		/// <summary>
		/// 18
		/// </summary>
		const int B_DBHGROWTHTRANSLATIONCODE=18;
		/// <summary>
		/// 19
		/// </summary>
		const int B_DBHGROWTHMEASUREMENTPERIOD=19;
		/// <summary>
		/// 20
		/// </summary>
		const int B_HTGROWTHTRANSLATIONCODE=20;
		/// <summary>
		/// 21
		/// </summary>
		const int B_HTGROWTHMEASUREMENTPERIOD=21;
		/// <summary>
		/// 22
		/// </summary>
		const int B_MORTMEASUREMENTPERIOD=22;
		/// <summary>
		/// 23
		/// </summary>
		const int B_MAXBASALAREA=23;
		/// <summary>
		/// 24
		/// </summary>
		const int B_MAXSTANDDENSITYINDEX=24;
		/// <summary>
		/// 25
		/// </summary>
		const int B_SITEINDEXSPECIES=25;
		/// <summary>
		/// 26
		/// </summary>
		const int B_SITEINDEX=26;
		/// <summary>
		/// 27
		/// </summary>
		const int B_MODELTYPE=27;
		/// <summary>
		/// 28
		/// </summary>
		const int B_PHYSIOGRAPHICREGIONCODE=28;

        struct FVS_TREE_ID
        {
            static public string strVariant;
            static public int intPlotId;
            static public int intTreeId;
            static public string strPlotId_formatted;
            static public string strTreeId_formatted;
        }

        


        

		



		public fvs_input(string p_strProjDir,frmTherm p_frmTherm)
		{
			//
			// TODO: Add constructor logic here
			//
			
			m_DataSource = new Datasource();
			m_DataSource.LoadTableColumnNamesAndDataTypes=false;
			m_DataSource.LoadTableRecordCount=false;
			m_DataSource.m_strDataSourceMDBFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\project.mdb";
			m_DataSource.m_strDataSourceTableName = "datasource";
			m_DataSource.m_strScenarioId="";
			m_DataSource.populate_datasource_array();
			this.m_strProjDir=p_strProjDir.Trim();

			this.m_strPlotTable = this.m_DataSource.getValidDataSourceTableName("PLOT");
			this.m_strCondTable = this.m_DataSource.getValidDataSourceTableName("CONDITION");
			this.m_strTreeTable = this.m_DataSource.getValidDataSourceTableName("TREE");
			this.m_strTreeSpcTable = this.m_DataSource.getValidDataSourceTableName("TREE SPECIES");
			this.m_strSiteTreeTable = this.m_DataSource.getValidDataSourceTableName("SITE TREE");
			this.m_strFVSTreeSpcTable = this.m_DataSource.getValidDataSourceTableName("FVS TREE SPECIES");
			
			if (this.m_strPlotTable.Trim().Length == 0) 
			{
				MessageBox.Show("!!Could Not Locate Plot Table!!","FVS Input",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}
			if (this.m_strCondTable.Trim().Length == 0) 
			{
				MessageBox.Show("!!Could Not Locate Condition Table!!","FVS Input",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}
			if (this.m_strTreeTable.Trim().Length == 0)
			{
				MessageBox.Show("!!Could Not Locate Tree Table!!","FVS Input",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}
			
			if (this.m_strTreeSpcTable.Trim().Length == 0)
			{
				MessageBox.Show("!!Could Not Locate Tree Species Table!!","FVS Input",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}
			if (this.m_strSiteTreeTable.Trim().Length == 0)
			{
				MessageBox.Show("!!Could Not Locate Site Tree Table!!","FVS Input",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}
			if (this.m_strFVSTreeSpcTable.Trim().Length == 0)
			{
				MessageBox.Show("!!Could Not Locate FVS Tree Species Table!!","FVS Input",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}

			this.m_strTempMDBFile = this.m_DataSource.CreateMDBAndTableDataSourceLinks();
			this.m_dao = new dao_data_access();
			this.m_ado = new ado_data_access();
			this.m_strConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"","");
			this.m_frmTherm = p_frmTherm;

			/*****************************************************************
			 **create a table structure to save biosum_cond_id, idb_alltree_id,
			 **tree cn,fvs_tree_id,fvs_filename
			 *****************************************************************/
			System.Data.DataTable p_dtSchema = this.m_ado.getTableSchema(this.m_strConn,"select biosum_cond_id,idb_alltree_id,fvs_tree_id,biosum_cond_id AS fvs_filename from " + this.m_strTreeTable + ";");
			this.m_strFVSTreeIdIDBWorkTable = this.m_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"fvs_tree_id_update_idb_work_table",p_dtSchema,true);
			p_dtSchema.Clear();
			p_dtSchema = this.m_ado.getTableSchema(this.m_strConn,"select biosum_cond_id,cn,fvs_tree_id,biosum_cond_id AS fvs_filename from " + this.m_strTreeTable + ";");
			
			this.m_strFVSTreeIdFIAWorkTable = this.m_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"fvs_tree_id_update_fia_work_table",p_dtSchema,true);

			m_strSlfALineArray = new string[5];
			m_strSlfBLineArray = new string[29];
			

		}
		~fvs_input()
		{
			try
			{
				if (this.m_ado.m_OleDbDataAdapter != null)
				{
					this.m_ado.m_OleDbDataAdapter.Dispose();
					this.m_ado.m_OleDbDataAdapter = null;
				}
				if (this.m_ado.m_OleDbConnection != null)
				{
					this.m_ado.m_OleDbConnection.Close();
					this.m_ado.m_OleDbConnection.Dispose();
					this.m_ado.m_OleDbConnection=null;
				}
				
			}
			catch
			{
			}
			this.m_ado = null;
			this.m_dao = null;
			this.m_DataSource = null;
		}
		public void Start(string p_strFVSInDir,string p_strVariant)
		{
            if (frmMain.g_bDebug)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "*****START*****" + System.DateTime.Now.ToString() + "\r\n");
            
            this.m_intError=0;
			this.m_strInDir = p_strFVSInDir.Trim() + "\\" + p_strVariant.Trim();
			this.m_strVariant = p_strVariant.Trim();
			CheckDir();
			DeleteFiles();
			if (this.m_intError!=0) return;
			InitializeFields();
			if (this.m_intError != 0) return;
			CreateLOC();
			if (this.m_intError != 0) return;
			CreateSLF();
			if (this.m_intError !=0) return;
			CreateFVS();

            if (frmMain.g_bDebug)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "*****END*****" + System.DateTime.Now.ToString() + "\r\n");

		}
		private void CreateLOC()
		{
			try
			{

				System.IO.FileStream p_fs = new System.IO.FileStream(this.m_strInDir + "\\" + this.m_strVariant + ".loc", System.IO.FileMode.Create, 
					System.IO.FileAccess.Write);
				System.IO.StreamWriter p_sw = new System.IO.StreamWriter(p_fs);
				p_sw.WriteLine("{0,-2}{1,-27}{2,-11}{3,-35}{4,-2}{5,-2}","A",'"' + this.m_strVariant + '"',"@",this.m_strVariant + ".slf","@","@");
				p_sw.Close();
				p_fs.Close();
				p_sw=null;
				p_fs=null;
			}
			catch (Exception e)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - fvs_input:CreateLOC  \n" + 
					"Err Msg - " + e.Message.ToString().Trim(),
					"FVS Input",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}


        
		}

		private void CreateSLF()
		{
			try
			{
				int x;
				int y;

				this.m_ado.OpenConnection(this.m_strConn);

				fvs_input.site_index oSiteIndex = new site_index();
                oSiteIndex.ado_data_access = m_ado;
				oSiteIndex.CondTable = this.m_strCondTable;
				oSiteIndex.PlotTable = this.m_strPlotTable;
				oSiteIndex.TreeTable = this.m_strTreeTable;
				oSiteIndex.SiteTreeTable = this.m_strSiteTreeTable;
				oSiteIndex.TreeSpeciesTable = this.m_strTreeSpcTable;
				oSiteIndex.FVSTreeSpeciesTable = this.m_strFVSTreeSpcTable;
                oSiteIndex.SiteIndexEquations = LoadSiteIndexEquations(this.m_strVariant.Trim().ToUpper());
                oSiteIndex.DebugFile = this.m_strDebugFile;

				this.m_ado.m_strSQL = "SELECT p.biosum_plot_id, c.biosum_cond_id, p.statecd ," + 
					"p.countycd, p.plot, p.fvs_variant, p.measyear," + 
					"c.adforcd,p.elev,c.condid, c.habtypcd1," + 
					"c.stdage,c.slope,c.aspect,c.ground_land_class_pnw," + 
					"c.sisp,p.lat,p.lon,p.idb_plot_id,c.adforcd,c.habtypcd1, " +
                    "p.elev,c.landclcd,c.ba_ft2_ac,c.habtypcd1 " + 
					"FROM " + this.m_strCondTable + " c," + 
					this.m_strPlotTable + " p " + 
					"WHERE p.biosum_plot_id = c.biosum_plot_id AND " + 
                    "c.landclcd=1 AND " + 
					"ucase(trim(p.fvs_variant)) = '" + this.m_strVariant.Trim().ToUpper() + "';";


				this.m_ado.CreateDataSet(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL,"slf");

				if (this.m_ado.m_DataSet.Tables["slf"].Rows.Count == 0)
				{
					this.m_intError = -1;
					MessageBox.Show("!!No SLF Records To Process!!", "FVS Input",
						System.Windows.Forms.MessageBoxButtons.OK,
						System.Windows.Forms.MessageBoxIcon.Exclamation);
					this.m_ado.m_DataSet.Clear();
					this.m_ado.m_DataSet.Dispose();
					this.m_ado.m_OleDbConnection.Close();
					this.m_ado.m_OleDbConnection.Dispose();
					return;

				}
				string strPlot="";
				int intStdAge=0;
				System.IO.FileStream p_fs = new System.IO.FileStream(this.m_strInDir + "\\" + this.m_strVariant + ".slf", System.IO.FileMode.Create, 
					System.IO.FileAccess.Write);
				System.IO.StreamWriter p_sw = new System.IO.StreamWriter(p_fs);
                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text", "Building SLF File For Variant " + this.m_strVariant);
                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Maximum", this.m_ado.m_DataSet.Tables["slf"].Rows.Count);
				this.m_dt = this.m_ado.m_DataSet.Tables["slf"];

				for (x=0; x<= this.m_dt.Rows.Count-1;x++)
				{
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", x);

				
					
					strPlot="";
					intStdAge=0;

					for (y=0;y<=this.m_strSlfALineArray.Length-1;y++)
					{
						this.m_strSlfALineArray[y]="";
					}

					for (y=0;y<=this.m_strSlfBLineArray.Length-1;y++)
					{
						this.m_strSlfBLineArray[y]="";
					}

					m_strSlfALineArray[fvs_input.A_RECORDTYPE]="A";
					m_strSlfALineArray[fvs_input.A_STANDID]=this.m_dt.Rows[x]["biosum_cond_id"].ToString().Trim();
					m_strSlfBLineArray[fvs_input.B_RECORDTYPE]="B";
					m_strSlfBLineArray[fvs_input.B_STANDID]=this.m_dt.Rows[x]["biosum_cond_id"].ToString().Trim();
					
					//process fields that are different for idb and fiadb
					if (this.m_dt.Rows[x]["biosum_plot_id"].ToString().Substring(0,1) == "1")
					{
						m_strSlfALineArray[fvs_input.A_FVSTREEFILENAME] = this.m_dt.Rows[x]["statecd"].ToString().PadLeft(2,'0');
						strPlot= this.m_dt.Rows[x]["plot"].ToString().PadLeft(5,'0');
						this.m_strSlfBLineArray[fvs_input.B_INVYR] = this.m_dt.Rows[x]["measyear"].ToString().Trim();
						
						//stand age
						if (this.m_dt.Rows[x]["stdage"] == System.DBNull.Value)
						{
							intStdAge=0;
						}
						else
						{
							if (this.m_dt.Rows[x]["stdage"].ToString().Trim().Length > 0)
							{
								intStdAge = Convert.ToInt32(this.m_dt.Rows[x]["stdage"]);
							}
							else
							{
								intStdAge=0;
							}
						}
					}
					else
					{
						m_strSlfALineArray[fvs_input.A_FVSTREEFILENAME] = this.IDB_getStateCd(this.m_dt.Rows[x]["statecd"].ToString().Trim());
						strPlot= this.m_dt.Rows[x]["idb_plot_id"].ToString().Trim();
						strPlot= strPlot.PadLeft(5,'0');
						/*******************************************************
						 **a few of the idb records do not have a 
						 **value in the measured year field. Get the year
						 **of the inventory as the year measured
						 *******************************************************/
						if (this.m_dt.Rows[x]["measyear"] == System.DBNull.Value)
						{
							this.m_strSlfBLineArray[fvs_input.B_INVYR] = IDB_getInvYear(this.m_dt.Rows[x]["biosum_plot_id"].ToString().Substring(1,4));
						}
						else
						{
							this.m_strSlfBLineArray[fvs_input.B_INVYR] = this.m_dt.Rows[x]["measyear"].ToString().Trim();
							if (m_strSlfBLineArray[fvs_input.B_INVYR].Length ==0)
							{
								this.m_strSlfBLineArray[fvs_input.B_INVYR] = IDB_getInvYear(this.m_dt.Rows[x]["biosum_plot_id"].ToString().Substring(1,4));
							}
							if (this.m_strSlfBLineArray[fvs_input.B_INVYR]=="1197")
							{
								this.m_strSlfBLineArray[fvs_input.B_INVYR]="1997";
							}
						}

						//stand age
						if (this.m_dt.Rows[x]["stdage"] == System.DBNull.Value)
						{
							intStdAge=0;
						}
						else
						{
							if (this.m_dt.Rows[x]["stdage"].ToString().Trim().Length > 0)
							{
								intStdAge = this.IDB_getStandAge(Convert.ToInt32(this.m_dt.Rows[x]["stdage"]));
							}
							else
							{
								intStdAge=0;
							}

						}
					}
					m_strSlfALineArray[fvs_input.A_FVSTREEFILENAME]+=strPlot + this.m_dt.Rows[x]["condid"].ToString().Trim() + ".fvs"; 
					m_strSlfALineArray[fvs_input.A_SAMPLEPOINTSITEDATAFLAG]="NoPointData  ";
					m_strSlfALineArray[fvs_input.A_FVSVARIANTSUSED] = m_strVariant;
					p_sw.WriteLine("{0,-2}{1,-27}{2,-102}{3,-13}{4,-5}{5,-2}",
						this.m_strSlfALineArray[fvs_input.A_RECORDTYPE],
						this.m_strSlfALineArray[fvs_input.A_STANDID],
						this.m_strSlfALineArray[fvs_input.A_FVSTREEFILENAME],
						this.m_strSlfALineArray[fvs_input.A_SAMPLEPOINTSITEDATAFLAG],
						this.m_strSlfALineArray[fvs_input.A_FVSVARIANTSUSED]," @");

					


					//latitude (fvs field 4)
					if (this.m_dt.Rows[x]["lat"] == System.DBNull.Value)
					{
						this.m_strSlfBLineArray[fvs_input.B_LAT] = "@";
					}
					else
					{
						this.m_strSlfBLineArray[fvs_input.B_LAT] = this.m_dt.Rows[x]["lat"].ToString().Trim();
						if (this.m_strSlfBLineArray[fvs_input.B_LAT].Length ==0)
						{
							this.m_strSlfBLineArray[fvs_input.B_LAT]="@";							
						}
						else
						{
							if (this.m_strSlfBLineArray[fvs_input.B_LAT] == ".") 
								this.m_strSlfBLineArray[fvs_input.B_LAT] = "@";
						}
					}
					this.m_strSlfBLineArray[fvs_input.B_LAT] = 
						this.m_strSlfBLineArray[fvs_input.B_LAT].PadRight(18,' ');

					//longitude (fvs field 5)
					if (this.m_dt.Rows[x]["lon"] == System.DBNull.Value)
					{
						this.m_strSlfBLineArray[fvs_input.B_LON] = "@";
					}
					else
					{
						this.m_strSlfBLineArray[fvs_input.B_LON] = this.m_dt.Rows[x]["lon"].ToString().Trim();
						if (this.m_strSlfBLineArray[fvs_input.B_LON].Length ==0)
						{
							this.m_strSlfBLineArray[fvs_input.B_LON]="@";							
						}
						else
						{
							if (this.m_strSlfBLineArray[fvs_input.B_LON] == ".") 
								this.m_strSlfBLineArray[fvs_input.B_LON] = "@";
						}
					}
					this.m_strSlfBLineArray[fvs_input.B_LON] = 
						this.m_strSlfBLineArray[fvs_input.B_LON].PadRight(18,' ');

					//USFS Region and Forest Code (location) (fvs field 6)
					if (this.m_dt.Rows[x]["adforcd"] == System.DBNull.Value)
					{
						this.m_strSlfBLineArray[fvs_input.B_FORESTCD] = "0";
					}
					else
					{
						this.m_strSlfBLineArray[fvs_input.B_FORESTCD] = this.m_dt.Rows[x]["adforcd"].ToString().Trim();
						if (this.m_strSlfBLineArray[fvs_input.B_FORESTCD].Length ==0)
						{
							this.m_strSlfBLineArray[fvs_input.B_FORESTCD]="0";							
						}
					}
					this.m_strSlfBLineArray[fvs_input.B_FORESTCD] = 
						this.m_strSlfBLineArray[fvs_input.B_FORESTCD].PadRight(5,' ');

					//Plant Association (habtypcd1) (fvs field 7)
					if (this.m_dt.Rows[x]["habtypcd1"] == System.DBNull.Value)
					{
						this.m_strSlfBLineArray[fvs_input.B_PLANTASSOC] = "@";
					}
					else
					{
						this.m_strSlfBLineArray[fvs_input.B_PLANTASSOC] = this.m_dt.Rows[x]["habtypcd1"].ToString().Trim();
						if (this.m_strSlfBLineArray[fvs_input.B_PLANTASSOC].Length ==0)
						{
							this.m_strSlfBLineArray[fvs_input.B_PLANTASSOC]="@";							
						}
						if (this.m_strSlfBLineArray[fvs_input.B_PLANTASSOC] == "n.a.")
							this.m_strSlfBLineArray[fvs_input.B_PLANTASSOC] = "@";
					}
					this.m_strSlfBLineArray[fvs_input.B_PLANTASSOC] = 
						this.m_strSlfBLineArray[fvs_input.B_PLANTASSOC].PadRight(10,' ');

					//Stand Year Of Origin (fvs field 8)
					this.m_strSlfBLineArray[fvs_input.B_STANDYEAROFORIGIN]= 
						Convert.ToString(Convert.ToInt32(this.m_strSlfBLineArray[fvs_input.B_INVYR]) - intStdAge);


					//aspect in degrees (fvs field 9)
					if (this.m_dt.Rows[x]["aspect"] == System.DBNull.Value)
					{
						this.m_strSlfBLineArray[fvs_input.B_ASPECT] = "@";
					}
					else
					{
						this.m_strSlfBLineArray[fvs_input.B_ASPECT] = this.m_dt.Rows[x]["Aspect"].ToString().Trim();
						if (this.m_strSlfBLineArray[fvs_input.B_ASPECT].Length ==0)
						{
							this.m_strSlfBLineArray[fvs_input.B_ASPECT]="@";							
						}
						if (this.m_strSlfBLineArray[fvs_input.B_ASPECT] == "-1") 
							this.m_strSlfBLineArray[fvs_input.B_ASPECT] = "@";
					}
					this.m_strSlfBLineArray[fvs_input.B_ASPECT] = 
						this.m_strSlfBLineArray[fvs_input.B_ASPECT].PadRight(4,' ');

					//slope percent (fvs field 10)
					if (this.m_dt.Rows[x]["elev"] == System.DBNull.Value)
					{
						this.m_strSlfBLineArray[fvs_input.B_SLOPEPERCENT] = "@";
					}
					else
					{
						m_strSlfBLineArray[fvs_input.B_SLOPEPERCENT] = this.m_dt.Rows[x]["Slope"].ToString().Trim();
						if (m_strSlfBLineArray[fvs_input.B_SLOPEPERCENT].Length ==0)
						{
							m_strSlfBLineArray[fvs_input.B_SLOPEPERCENT]="@";							
						}
						if (m_strSlfBLineArray[fvs_input.B_SLOPEPERCENT] == "-1")
							m_strSlfBLineArray[fvs_input.B_SLOPEPERCENT] = "@";
					}
					m_strSlfBLineArray[fvs_input.B_SLOPEPERCENT] = m_strSlfBLineArray[fvs_input.B_SLOPEPERCENT].PadRight(4,' ');

					//elevation in 100's of feet (fvs field 11)
					if (this.m_dt.Rows[x]["elev"] == System.DBNull.Value)
					{
						this.m_strSlfBLineArray[fvs_input.B_ELEV] = "@";
					}
					else
					{
						m_strSlfBLineArray[fvs_input.B_ELEV] = this.m_dt.Rows[x]["elev"].ToString().Trim();
						if (m_strSlfBLineArray[fvs_input.B_ELEV].Length ==0)
						{
							m_strSlfBLineArray[fvs_input.B_ELEV]="@";							
						}
						if (m_strSlfBLineArray[fvs_input.B_ELEV] != "@")
						{
							if (Convert.ToInt32(m_strSlfBLineArray[fvs_input.B_ELEV]) > 0)
							{
								m_strSlfBLineArray[fvs_input.B_ELEV] = 
									Convert.ToString(Convert.ToInt32(m_strSlfBLineArray[fvs_input.B_ELEV]) / 100);
								m_strSlfBLineArray[fvs_input.B_ELEV] = 
									Convert.ToString(Math.Round(Convert.ToDouble(m_strSlfBLineArray[fvs_input.B_ELEV])));
							}

						}
					}
					m_strSlfBLineArray[fvs_input.B_ELEV]= 
						m_strSlfBLineArray[fvs_input.B_ELEV].PadRight(4,' ');

					this.m_strSlfBLineArray[fvs_input.B_BASALAREAFACTOR]="0";

					this.m_strSlfBLineArray[fvs_input.B_INVERSESMALLTREE]="1";

					this.m_strSlfBLineArray[fvs_input.B_DBHBREAKPOINT]="999";

					this.m_strSlfBLineArray[fvs_input.B_NUMBEROFPLOTSPERFILE]="1";

					m_strSlfBLineArray[fvs_input.B_NUMBEROFNONSTOCKABLEPLOTS]="0";

					m_strSlfBLineArray[fvs_input.B_STANDSAMPLINGWEIGHT]="1";

					//land class code (proportion of stand considered stockable) (fvs field 12)
					if (this.m_dt.Rows[x]["landclcd"] == System.DBNull.Value)
					{
						this.m_strSlfBLineArray[fvs_input.B_STOCKABLEPERCENT] = "0";
					}
					else
					{
						m_strSlfBLineArray[fvs_input.B_STOCKABLEPERCENT] = 
							this.m_dt.Rows[x]["landclcd"].ToString().Trim();
						if (m_strSlfBLineArray[fvs_input.B_STOCKABLEPERCENT].Length ==0)
						{
							m_strSlfBLineArray[fvs_input.B_STOCKABLEPERCENT]="0";							
						}
						else if (m_strSlfBLineArray[fvs_input.B_STOCKABLEPERCENT] != "1")
							m_strSlfBLineArray[fvs_input.B_STOCKABLEPERCENT] = "0";
					}
					m_strSlfBLineArray[fvs_input.B_STOCKABLEPERCENT] = 
						m_strSlfBLineArray[fvs_input.B_STOCKABLEPERCENT].PadRight(2,' ');


					m_strSlfBLineArray[fvs_input.B_DBHGROWTHTRANSLATIONCODE]="1";

					m_strSlfBLineArray[fvs_input.B_DBHGROWTHMEASUREMENTPERIOD]="10";

					m_strSlfBLineArray[fvs_input.B_HTGROWTHTRANSLATIONCODE]="1";

					m_strSlfBLineArray[fvs_input.B_HTGROWTHMEASUREMENTPERIOD]="5";

					m_strSlfBLineArray[fvs_input.B_MORTMEASUREMENTPERIOD]="5";

					m_strSlfBLineArray[fvs_input.B_MAXBASALAREA]="@";

					m_strSlfBLineArray[fvs_input.B_MAXSTANDDENSITYINDEX]="@";

					oSiteIndex.getSiteIndex(m_dt.Rows[x]);

                    m_strSlfBLineArray[fvs_input.B_SITEINDEXSPECIES] = oSiteIndex.SiteIndexSpeciesAlphaCode;

					m_strSlfBLineArray[fvs_input.B_SITEINDEX]=oSiteIndex.SiteIndex;

					m_strSlfBLineArray[fvs_input.B_MODELTYPE]="@";

					m_strSlfBLineArray[fvs_input.B_PHYSIOGRAPHICREGIONCODE]="@";

					p_sw.WriteLine("{0,-2}{1,-27}{2,-5}{3,-5}{4,-5}{5,-5}{6,-11}" + 
						"{7,-5}{8,-4}{9,-4}{10,-4}{11,-2}{12,-2}{13,-4}{14,-2}{15,-2}" + 
						"{16,-2}{17,-2}{18,-2}{19,-3}{20,-2}{21,-2}{22,-2}{23,-2}" + 
						"{24,-2}{25,-4}{26,-17}{27,-2}{28,-2}{29,-2}",
						m_strSlfBLineArray[fvs_input.B_RECORDTYPE],
						m_strSlfBLineArray[fvs_input.B_STANDID],
						m_strSlfBLineArray[fvs_input.B_INVYR],
						m_strSlfBLineArray[fvs_input.B_LAT],
						m_strSlfBLineArray[fvs_input.B_LON],
						m_strSlfBLineArray[fvs_input.B_FORESTCD],
						m_strSlfBLineArray[fvs_input.B_PLANTASSOC],
						m_strSlfBLineArray[fvs_input.B_STANDYEAROFORIGIN],
						m_strSlfBLineArray[fvs_input.B_ASPECT],
						m_strSlfBLineArray[fvs_input.B_SLOPEPERCENT],
						m_strSlfBLineArray[fvs_input.B_ELEV],
						m_strSlfBLineArray[fvs_input.B_BASALAREAFACTOR],
						m_strSlfBLineArray[fvs_input.B_INVERSESMALLTREE],
						m_strSlfBLineArray[fvs_input.B_DBHBREAKPOINT],
						m_strSlfBLineArray[fvs_input.B_NUMBEROFPLOTSPERFILE],
						m_strSlfBLineArray[fvs_input.B_NUMBEROFNONSTOCKABLEPLOTS],
						m_strSlfBLineArray[fvs_input.B_STANDSAMPLINGWEIGHT],
						m_strSlfBLineArray[fvs_input.B_STOCKABLEPERCENT],
						m_strSlfBLineArray[fvs_input.B_DBHGROWTHTRANSLATIONCODE],
						m_strSlfBLineArray[fvs_input.B_DBHGROWTHMEASUREMENTPERIOD],
						m_strSlfBLineArray[fvs_input.B_HTGROWTHTRANSLATIONCODE],
						m_strSlfBLineArray[fvs_input.B_HTGROWTHMEASUREMENTPERIOD],
						m_strSlfBLineArray[fvs_input.B_MORTMEASUREMENTPERIOD],
						m_strSlfBLineArray[fvs_input.B_MAXBASALAREA],
						m_strSlfBLineArray[fvs_input.B_MAXSTANDDENSITYINDEX],
						m_strSlfBLineArray[fvs_input.B_SITEINDEXSPECIES],
						m_strSlfBLineArray[fvs_input.B_SITEINDEX],
						m_strSlfBLineArray[fvs_input.B_MODELTYPE],
						m_strSlfBLineArray[fvs_input.B_PHYSIOGRAPHICREGIONCODE],
						"@");




					//B 00525C03          1997 42  -125 0       @         1932 315 15  1   0      1      999 1    0    1       1    1 10 1 5  5  @   @   DF  97        @ @  @    @

				}
				//close the connection to the temp mdb file
				this.m_ado.m_OleDbConnection.Close();
				while (m_ado.m_OleDbConnection.State == System.Data.ConnectionState.Open)
					System.Threading.Thread.Sleep(1000);
				m_ado.m_OleDbConnection.Dispose();
				m_ado.m_OleDbConnection=null;
				oSiteIndex = null;
				p_sw.Close();
				p_fs.Close();
				p_sw=null;
				p_fs=null;
			}
			catch (Exception e)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - fvs_input:CreateSLF  \n" + 
					"Err Msg - " + e.Message.ToString().Trim(),
					"FVS Input",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}
			finally
			{
				if (m_ado.m_OleDbConnection != null)
				{
					while (m_ado.m_OleDbConnection.State == System.Data.ConnectionState.Open)
						System.Threading.Thread.Sleep(1000);
					m_ado.m_OleDbConnection.Dispose();
					m_ado.m_OleDbConnection=null;
				}
				
			}


        
		}

		/// <summary>
		/// create fvs tree input
		/// </summary>
		public void CreateFVS()
		{
			int x;
			int y;
			int z;
			

			const int LIVE_TREE = 1;
			const int MORT_TREE = 2;
			const int FIA_SPCD = 0;
			const int CONVERT_TO_FIA_SPCD = 1;
			const int FVS_SPECIES=2;
			string str="";
			int intPlotCnt=0;
			int intTreeCnt=0;
			string strPlotId="";
			string strTreeId="";
			string strSQLWhereExp="";
			string strCurRecordId="";			
			string strRecordId;					//biosum condition id
			string strStateCd;					//state code
			string strPlotCd;					//plot code
			string strCondId;					//condition id
			string strCN="";					//fiadb unique id for a tree record
			string strIDBAlltreeId="";          //idb unique id for a tree record
			
			string strFVSDirAndFile="";	
			string strFVSFile="";
			string strFVSTreeId="";
			string strTpa="";					//trees per acre. 1 decimal place
			string strTpa2="";					//trees per acre. 2 decimal places
			double dblTpa=0;					//trees per acre. 1 decimal place
			double dblTpa2=0;                   //trees per acre. 2 decimal places
			bool bUsePlotCnt = true;
			bool bSkip;
			int intDec;							
			int intFVSPlotId=0;					//fvs assigned plot id
			string strHistCd="";				//history code
			string strFvsCvtSpCd="";				//converted species code
			
			string strDbh="";					//diameter
			double dblDbh;						//diameter
            string strDbhInc10yr="";            //diameter 10 year increment
            double dblDbhInc10yr;               //diameter 10 year increment
			string strHtCd;						//height code
			string strHt;						//height in feet
			string strActualHt;					//actual height
			string strCr;                       //crown ratio
			System.IO.FileStream p_fs;
			System.IO.StreamWriter p_sw;
			string[,] strSpCd;
			strSpCd = new string[700,3];
			int intSpcCvtCnt;
			int intSpcRec=0;
            bool bNewPlot = false;
            int intMaxPlotId = -1;
            int intMaxTreeId = -1;

			string [,] strFvsSpCd;
			strFvsSpCd = new string[700,2];
			
			

			string strDmgAg1="";
			string strDmgAg2="";
			string strDmgAg3="";
			string strDmgSv1="";
			string strDmgSv2="";
			string strDmgSv3="";
			
		    double    dblCullBf=0;
			
			string strValueClass="";
			string strMistClCd = "";
			bool bBrokenTop;
            int intFvsTreeIdRecordCount = 0;


		

			p_fs = new System.IO.FileStream(this.m_strProjDir + "\\fvs\\data\\null.txt", System.IO.FileMode.Create, 
				System.IO.FileAccess.Write);

			p_sw = new System.IO.StreamWriter(p_fs);


			try
			{
 
				this.m_ado.OpenConnection(this.m_strConn);

				//load up the spcd conversion table into an array
                /**************************************************************
                 **Some FIA species are not recognized by FVS within the 
                 **FVS variant context. FVS will automatically convert those
                 **unrecognized species to either 298, 998 or other generic
                 ***species. The tree species table provides a conversion matrix
                 **to avoid FVS changing the species codes. FVS unrecognized
                 **species can be converted to a species code within the 
                 **family genus.
                 **************************************************************/
				this.m_ado.m_strSQL = "SELECT spcd,fvs_input_spcd,fvs_species " + 
					"FROM " + this.m_strTreeSpcTable + " " + 
					"WHERE trim(ucase(fvs_variant)) = '" + this.m_strVariant.Trim() + "';";
				this.m_ado.SqlQueryReader(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);

				x=0;
				while (this.m_ado.m_OleDbDataReader.Read())
				{
					strSpCd[x,FIA_SPCD] = this.m_ado.m_OleDbDataReader["spcd"].ToString().Trim();
					if (this.m_ado.m_OleDbDataReader["fvs_input_spcd"] == System.DBNull.Value)
					{
						strSpCd[x,CONVERT_TO_FIA_SPCD] = strSpCd[x,FIA_SPCD];
					}
					else
					{
						strSpCd[x,CONVERT_TO_FIA_SPCD] = this.m_ado.m_OleDbDataReader["fvs_input_spcd"].ToString().Trim();
					}
					if (this.m_ado.m_OleDbDataReader["fvs_species"] == System.DBNull.Value ||
						this.m_ado.m_OleDbDataReader["fvs_species"].ToString().Trim().Length == 0)
					{
						strSpCd[x,FVS_SPECIES] = "";
					}
					else
					{
						strSpCd[x,FVS_SPECIES] = this.m_ado.m_OleDbDataReader["fvs_species"].ToString().Trim();
					}
					x++;
				}
				this.m_ado.m_OleDbDataReader.Close();
				intSpcCvtCnt=x;	

                //load up any previously assigned fvs_tree_id's into a table
                if (m_ado.TableExist(m_ado.m_OleDbConnection,"fvs_tree_id_work_table"))
                    m_ado.SqlNonQuery(m_ado.m_OleDbConnection, "DROP TABLE fvs_tree_id_work_table");

                m_ado.m_strSQL = Tables.FVS.CreateFVSTreeIdWorkTableSQL("fvs_tree_id_work_table");
                m_ado.SqlNonQuery(m_ado.m_OleDbConnection, m_ado.m_strSQL);

                m_ado.m_strSQL = "INSERT INTO fvs_tree_id_work_table " + 
                                 "SELECT biosum_cond_id, " +
                                        "fvs_tree_id," + 
                                        "MID(FVS_TREE_ID,1,2) AS fvs_variant," +
                                        "MID(FVS_TREE_ID,3,4) AS plotid," +
                                        "MID(FVS_TREE_ID,7,3) AS treeid," +
                                        "MID(FVS_TREE_ID,3,10) AS plottreeid " +
                                 "FROM " + this.m_strTreeTable + " " +
                                 "WHERE FVS_TREE_ID IS NOT NULL AND " +
                                       "MID(FVS_TREE_ID,1,2)='" + m_strVariant.Trim() + "'";
                m_ado.SqlNonQuery(m_ado.m_OleDbConnection, m_ado.m_strSQL);

                /*****************************************************************
                 **If intFvsTreeIdRecordCount
                 **is zero then no tree record has
                 **a fvs_tree_id value and input files will be created within the
                 **code and will not reference the fvs_tree_id_work_table.
                 **
                 **If intFvsTreeIdRecordCount
                 **is greater than zero then fvs_tree_id will be created using
                 **fvs_tree_id_work_table
                 *****************************************************************/
                intFvsTreeIdRecordCount = (int)m_ado.getRecordCount(
                                                m_ado.m_OleDbConnection,
                                                "SELECT COUNT(*) FROM fvs_tree_id_work_table",
                                                "fvs_tree_id_work_table");

				System.Data.DataTable p_dt = new System.Data.DataTable("id");
				p_dt.Columns.Add("biosum_cond_id",typeof(string));
				p_dt.Columns.Add("fvs_plot_id",typeof(int));
				p_dt.Columns.Add("fvs_tree_id",typeof(int));
				p_dt.Columns.Add("idb_alltree_id",typeof(int));
				p_dt.Columns.Add("cn",typeof(string));

				for (x=LIVE_TREE;x<=MORT_TREE;x++)
				{
					if (x==LIVE_TREE)
					{
						
						this.m_ado.m_strSQL = "SELECT c.biosum_cond_id, p.statecd, " + 
							"p.countycd, p.plot,p.measyear,c.adforcd, " + 
							"p.elev,c.condid, c.habtypcd1,c.stdage,c.slope,c.aspect," + 
							"c.ground_land_class_pnw,c.landclcd, " + 
							"c.sisp,p.idb_plot_id,t.spcd, t.condid as tree_condid," + 
							"t.idb_alltree_id,t.cn,t.dia,t.ht,t.htcd,t.actualht,t.cr,t.tpacurr, " + 
							"t.agentcd,t.damtyp1,t.damsev1," + 
							"t.damtyp2,t.damsev2,t.idb_dmg_agent3_cd,t.idb_severity3_cd," + 
							"t.cullbf,t.mist_cl_cd,t.fvs_dmg_ag1,t.fvs_dmg_sv1," + 
							"t.fvs_dmg_ag2,t.fvs_dmg_sv2,t.fvs_dmg_ag3,t.fvs_dmg_sv3,t.inc10yr,t.fvs_tree_id " + 
							"FROM " + this.m_strCondTable + " c," + 
							this.m_strPlotTable + " p," + 
							this.m_strTreeTable + " t ";
							
						



						strSQLWhereExp = "WHERE ucase(trim(p.fvs_variant)) = '" + this.m_strVariant.Trim().ToUpper() + "' AND " + 
							"p.biosum_plot_id = c.biosum_plot_id AND " + 
							"t.biosum_cond_id = c.biosum_cond_id AND t.statuscd=1 AND c.landclcd=1";

					}
					else
					{
						this.m_ado.m_strSQL = "SELECT c.biosum_cond_id, p.statecd, " + 
							"p.countycd, p.plot,p.measyear,c.adforcd, " + 
							"p.elev,c.condid, c.habtypcd1,c.stdage,c.slope,c.aspect," + 
							"c.ground_land_class_pnw,c.landclcd, " + 
							"c.sisp,p.idb_plot_id,t.spcd, t.condid as tree_condid," + 
							"t.idb_alltree_id,t.cn,t.dia,t.ht,t.htcd,t.actualht,t.cr,t.tpacurr,t.inc10yr,t.fvs_tree_id " + 
							"FROM " + this.m_strCondTable + " c," + 
							this.m_strPlotTable + " p," + 
							this.m_strTreeTable + " t ";

						strSQLWhereExp = "WHERE ucase(trim(p.fvs_variant)) = '" + this.m_strVariant.Trim().ToUpper() + "' AND " + 
							"p.biosum_plot_id = c.biosum_plot_id AND " + 
							"t.biosum_cond_id = c.biosum_cond_id AND t.statuscd <> 1 AND c.landclcd=1";
					}
					this.m_ado.m_strSQL += " " + strSQLWhereExp + " ORDER BY c.biosum_cond_id;";

					//initialize the dataset
					this.m_ado.m_DataSet.Clear();
					//initialize the datatable
					this.m_dt.Clear();

					//create the dataset
					this.m_ado.CreateDataSet(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL,"fvs");
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Minimum", 0);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Maximum", this.m_ado.m_DataSet.Tables["fvs"].Rows.Count);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", 0);
					if (x == LIVE_TREE)
					{
                        frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text", "Building FVS File(s) - Live Trees");
					}
					else
					{
                        frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text", "Building FVS File(s) - Mortality Trees");
					}
                    frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Refresh");

					if (this.m_ado.m_DataSet.Tables["fvs"].Rows.Count > 0)
					{
						this.m_dt = this.m_ado.m_DataSet.Tables["fvs"];
                    
						//process each tree record
						for (y=0;y<=this.m_dt.Rows.Count-1;y++)
						{
							//initialize record variables
							strTreeId="";			   //FVS assigned value
							strCN = "";				   //FIADB tree unique identifier
							strIDBAlltreeId="";        //IDB tree unique identifier
							dblTpa=0;
							dblTpa2=0;
							strTpa="";				   //trees per acre
							strTpa2="";
							strHistCd="";
							strFvsCvtSpCd="";
							strDbh="";
							dblDbh=0;
                            strDbhInc10yr = "";
                            dblDbhInc10yr = 0;
							strHt="0";
							strHtCd="";
							strActualHt="0";
							strCr=" ";
							strDmgAg1="0";
							strDmgAg2="0";
							strDmgAg3="0";
							strDmgSv1="0";
							strDmgSv2="0";
							strDmgSv3="0";
							dblCullBf=0;
							strValueClass="3";
							strMistClCd = "";
							intSpcRec=-1;
							bBrokenTop=false;


							bSkip=false;
                            frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", y);

							//biosum_cond_id
							strRecordId = Convert.ToString(this.m_dt.Rows[y]["biosum_cond_id"]).Trim();

                            
							//statecd
							strStateCd = Convert.ToString(this.m_dt.Rows[y]["statecd"]).Trim();

							//check if this record is FIADB or IDB
							if (strRecordId.Substring(0,1) == "1")
							{
								//fiadb
								strPlotCd = Convert.ToString(this.m_dt.Rows[y]["plot"]).Trim();
							}
							else
							{
								//idb
								strPlotCd = Convert.ToString(this.m_dt.Rows[y]["idb_plot_id"]).Trim();
								strStateCd = this.IDB_getStateCd(strStateCd);
							}
						
							strCondId = Convert.ToString(this.m_dt.Rows[y]["condid"]).Trim();

                            if (this.m_dt.Rows[y]["fvs_tree_id"] != DBNull.Value &&
                                this.m_dt.Rows[y]["fvs_tree_id"].ToString().Trim().Length > 0)
                            {
                                FVS_TREE_ID.strVariant = m_dt.Rows[y]["FVS_TREE_ID"].ToString().Substring(0, 2);
                                FVS_TREE_ID.intPlotId = Convert.ToInt32(m_dt.Rows[y]["FVS_TREE_ID"].ToString().Substring(2, 4));
                                FVS_TREE_ID.strPlotId_formatted = m_dt.Rows[y]["FVS_TREE_ID"].ToString().Substring(2, 4);
                                FVS_TREE_ID.intTreeId = Convert.ToInt32(m_dt.Rows[y]["FVS_TREE_ID"].ToString().Substring(6, 3));
                                FVS_TREE_ID.strTreeId_formatted = m_dt.Rows[y]["FVS_TREE_ID"].ToString().Substring(6, 3);
                            }
                            else
                            {
                                
                                FVS_TREE_ID.intPlotId = -1;
                                FVS_TREE_ID.intTreeId = -1;
                                FVS_TREE_ID.strVariant = this.m_strVariant;
                                FVS_TREE_ID.strPlotId_formatted = "";
                                FVS_TREE_ID.strTreeId_formatted = "";

                            }
						

							//check if this is a new plot/condition
							if (strCurRecordId != strRecordId)
							{
								bUsePlotCnt=false;
								strCurRecordId = strRecordId;
								
								//initialize to the file to output fvs data
								strFVSDirAndFile= strStateCd.PadLeft(2,'0') + strPlotCd.PadLeft(5,'0') + strCondId + ".fvs";
								strFVSFile = strFVSDirAndFile;
								strFVSDirAndFile = this.m_strProjDir + "\\fvs\\data\\" + this.m_strVariant.Trim() + "\\" + strFVSDirAndFile;

								//close the old out fvs data file if it is open
								if (p_fs != null)
								{
									p_sw.Close();
									p_fs.Close();
								}

								//if the output fvs data file exists than append data
                                if (System.IO.File.Exists(strFVSDirAndFile) == true)
                                {
                                    if (intFvsTreeIdRecordCount == 0)
                                    {
                                        if (x == MORT_TREE)
                                        {
                                            //read the first line to get the previous FVS assigned plot id
                                            System.IO.StreamReader p_sr = new System.IO.StreamReader(strFVSDirAndFile);
                                            string strLine = p_sr.ReadLine();
                                            if (strLine != null && strLine.Trim().Length > 4)
                                            {
                                                intFVSPlotId = Convert.ToInt32(strLine.Substring(0, 4));
                                            }
                                            else
                                            {
                                                bUsePlotCnt = true;
                                                intPlotCnt++;
                                            }
                                            p_sr.Close();
                                            p_sr = null;
                                            p_fs = new System.IO.FileStream(strFVSDirAndFile, System.IO.FileMode.Append,
                                                System.IO.FileAccess.Write);

                                        }
                                        else
                                        {
                                            p_fs = new System.IO.FileStream(strFVSDirAndFile, System.IO.FileMode.Append,
                                                System.IO.FileAccess.Write);

                                            bUsePlotCnt = true;
                                            intPlotCnt++;
                                        }
                                    }
                                    else
                                    {
                                        if (x == MORT_TREE)
                                        {
                                            p_fs = new System.IO.FileStream(strFVSDirAndFile, System.IO.FileMode.Append,
                                                System.IO.FileAccess.Write);

                                        }
                                        else
                                        {
                                            p_fs = new System.IO.FileStream(strFVSDirAndFile, System.IO.FileMode.Append,
                                                System.IO.FileAccess.Write);
                                        }

                                    }

                                }
                                else
                                {
                                    p_fs = new System.IO.FileStream(strFVSDirAndFile, System.IO.FileMode.Create,
                                        System.IO.FileAccess.Write);
                                    if (intFvsTreeIdRecordCount == 0)
                                    {
                                        bUsePlotCnt = true;
                                        intPlotCnt++;
                                    }

                                }

                                if (intFvsTreeIdRecordCount == 0)
                                {
                                    //check if we use the incremented plot id count
                                    //or the previously fvs assigned plot id
                                    if (bUsePlotCnt == true) intFVSPlotId = intPlotCnt;

                                    strPlotId = Convert.ToString(intFVSPlotId).Trim();
                                    strPlotId = strPlotId.PadLeft(4, '0');
                                }
                               							
								//create a new instance of the stream writer class
								p_sw = new System.IO.StreamWriter(p_fs);

                                if (intFvsTreeIdRecordCount == 0)
                                {
                                    //live trees are assigned fvs tree ids from 0 to 499
                                    //mortality trees are assigned fvs tree ids from 500 to 999.
                                    //initialize the counts since this is a new plot/condition
                                    if (x == LIVE_TREE)
                                    {
                                        intTreeCnt = 0;
                                    }
                                    else
                                    {
                                        intTreeCnt = 499;
                                    }
                                }
								
							}
                            //check which tree id record to use
							if (strRecordId.Substring(0,1) == "1")
							{
								//FIADB
								strCN = Convert.ToString(this.m_dt.Rows[y]["cn"]).Trim();
							}
							else
							{
								//IDB
								strIDBAlltreeId = Convert.ToString(this.m_dt.Rows[y]["idb_alltree_id"]);
							}


                            if (intFvsTreeIdRecordCount == 0)
                            {
                                /********************************************************************************
                                **to get a unique tree id for each plot, live trees are numbered from 1 to 499
                                **and dead trees are numbered from 500 to 999
                                ************************ ********************************************************/
                                if (x == LIVE_TREE)
                                {
                                    if (intTreeCnt == 499)
                                    {
                                        intTreeCnt = 1;
                                    }
                                    else
                                        intTreeCnt++;
                                }
                                else
                                {
                                    if (intTreeCnt == 999)
                                    {
                                        intTreeCnt = 500;
                                    }
                                    else
                                        intTreeCnt++;
                                }

                                strTreeId = Convert.ToString(intTreeCnt).Trim();
                                strTreeId = strTreeId.PadLeft(3, '0');
                            }

							//tpa (trees per acre)
							if (this.m_dt.Rows[y]["tpacurr"] == System.DBNull.Value)
							{
								strTpa = "0";
							}
							else
							{
								strTpa = this.m_dt.Rows[y]["tpacurr"].ToString().Trim();
								strTpa2=strTpa;
								if (strTpa.Length ==0)
								{
									strTpa="0";							
								}
								
							}
							dblTpa = Convert.ToDouble(strTpa);
							dblTpa = Math.Round(dblTpa,1);
							
							//do not write this tree record if TPA <=0
							if (dblTpa <= 0) bSkip=true;

							if (bSkip==false)
							{

                                if (dblTpa > 9999.9)
                                {
                                    dblTpa = Math.Round(dblTpa, 0);
                                    strTpa = Convert.ToString(dblTpa);

                                }
                                else
                                {
                                    strTpa = Convert.ToString(dblTpa);

                                    intDec = strTpa.IndexOf(".", 0);


                                    if (intDec < 0)
                                    {
                                        strTpa = strTpa + ".0";
                                    }
                                }

                                strTpa = strTpa.PadLeft(6, ' ');


								//tree hist cd
								if (x==LIVE_TREE) strHistCd="1";
								else strHistCd="9";

								//convert FIA SpCd to  FVS SpCd
								if (this.m_dt.Rows[y]["spcd"] == System.DBNull.Value)
								{
									strFvsCvtSpCd="";
								}
								else
								{
									str = Convert.ToString(this.m_dt.Rows[y]["spcd"]).Trim();
									for (z=0;z<=intSpcCvtCnt-1;z++)
									{
										if (str  == strSpCd[z,FIA_SPCD]) 
										{
											strFvsCvtSpCd = strSpCd[z,CONVERT_TO_FIA_SPCD];
											intSpcRec = z;
											break;
										}
									}
								}
								strFvsCvtSpCd = strFvsCvtSpCd.PadLeft(3,'0');


								//dbh
								if (this.m_dt.Rows[y]["dia"] == System.DBNull.Value)
								{
									dblDbh = 0;
								}
								else
								{
									dblDbh = Convert.ToDouble(this.m_dt.Rows[y]["dia"]);
								}
								dblDbh = Math.Round(dblDbh,1);
								if (x==LIVE_TREE)
								{
									if (dblTpa > 25 && dblDbh <= 0) dblDbh = 0.1; //assume to be a seedling
								}
								if (dblDbh <= 0) bSkip=true;
								if (bSkip == false)
								{
									strDbh = Convert.ToString(dblDbh).Trim();
									intDec = strDbh.IndexOf(".",0);
                                    if (intDec < 0)
                                    {
                                        if (strDbh.Trim().Length == 3)
                                        {
                                            strDbh = strDbh + "0";
                                        }
                                        else if (strDbh.Trim().Length == 2)
                                        {
                                            strDbh = "0" + strDbh + "0";
                                        }
                                        else if (strDbh.Trim().Length == 1)
                                        {
                                            strDbh = "00" + strDbh + "0";
                                        }
                                    }
                                    else if (strDbh.Trim().Length == 5)
                                    {
                                        strDbh = strDbh.Replace(".", "");
                                    }
                                    else if (strDbh.Trim().Length == 4)
                                    {
                                        strDbh = "0" + strDbh.Replace(".", "");
                                    }
                                    else if (strDbh.Trim().Length == 3)
                                    {
                                        strDbh = "00" + strDbh.Replace(".", "");
                                    }
                                    else if (strDbh.Trim().Length == 2)
                                    {
                                        strDbh = "000" + strDbh.Replace(".", "");
                                    }
                                    //10 year increment
                                    if (this.m_dt.Rows[y]["inc10yr"] == System.DBNull.Value)
                                    {
                                        strDbhInc10yr = "   ";
                                    }
                                    else
                                    {
                                        dblDbhInc10yr = (double)Convert.ToDouble(this.m_dt.Rows[y]["inc10yr"]) * (double).1;
                                        dblDbhInc10yr = (double)Math.Round(dblDbhInc10yr, 1);
                                        strDbhInc10yr = Convert.ToString(dblDbhInc10yr).Trim();
                                        intDec = strDbhInc10yr.IndexOf(".", 0);
                                        if (intDec < 0)
                                        {
                                            if (strDbhInc10yr.Trim().Length == 2)
                                            {
                                                strDbhInc10yr = strDbhInc10yr + "0";
                                            }
                                            else if (strDbhInc10yr.Trim().Length == 1)
                                            {
                                                strDbhInc10yr = "0" + strDbhInc10yr + "0";
                                            }
                                        }
                                        else if (strDbhInc10yr.Trim().Length == 4)
                                        {
                                            strDbhInc10yr = strDbhInc10yr.Replace(".", "");
                                        }
                                        else if (strDbhInc10yr.Trim().Length == 3)
                                        {
                                            strDbhInc10yr = "0" + strDbhInc10yr.Replace(".", "");
                                        }
                                        else if (strDbhInc10yr.Trim().Length == 2)
                                        {
                                            strDbhInc10yr = "00" + strDbhInc10yr.Replace(".", "");
                                        }
                                        
                                       

                                    }
								    

                                    //note: did not process dbh increment because no growth estimates

                                    //height code
									if (this.m_dt.Rows[y]["htcd"] == System.DBNull.Value)
									{
										strHtCd = " ";
									}
									else
									{
										strHtCd= Convert.ToString(this.m_dt.Rows[y]["htcd"]);

									}

									//height in feet
									//if htcd has a value or pnw idb record then process ht value
									if (strHtCd=="1" || strHtCd=="2" || strHtCd=="3" || strRecordId.Substring(0,1) == "2") 
									{
										if (this.m_dt.Rows[y]["ht"] == System.DBNull.Value)
										{
											strHt = "0";
										}
										else
										{
											strHt= Convert.ToString(this.m_dt.Rows[y]["ht"]).Trim();
											if (strHt.Trim().Length==0) strHt="0";
											if (Convert.ToDouble(strHt) < 0) strHt="0";

										}

										//height to top kill (broken tops)
										//pnw idb does not have broken tops for live trees so by pass actual ht
										if (strRecordId.Substring(0,1) != "2")
										{
											if (this.m_dt.Rows[y]["actualht"] == System.DBNull.Value)
											{
												strActualHt="0";
											}
											else
											{
												strActualHt= Convert.ToString(this.m_dt.Rows[y]["actualht"]).Trim();
												if (strActualHt.Length ==0) strActualHt="0";
												if (Convert.ToDouble(strHt) <= Convert.ToDouble(strActualHt))
												{
													strActualHt="0";
												}
												else
												{
													bBrokenTop=true;
												}
											}
										}
									}

									//note: by pass ht increment because not doing growth estimates


									//compacted crown ratio
									if (x==LIVE_TREE)
									{
										//if have measured htcd or pnw idb record than process crown ratio
										if (strHtCd=="1" || strHtCd=="2" || strHtCd=="3" || strRecordId.Substring(0,1) == "2") 
										{
											if (this.m_dt.Rows[y]["cr"] != System.DBNull.Value)
											{
												strCr= Convert.ToString(this.m_dt.Rows[y]["cr"]).Trim();
												if (strCr.Length ==0) strCr="0";
												if (Convert.ToDouble(strCr) > 0)
												{
													double dblCr = Convert.ToDouble(strCr);
													if (dblCr <= 10) strCr = "1";
													else if (dblCr <= 20) strCr = "2";
													else if (dblCr <= 30) strCr = "3";
													else if (dblCr <= 40) strCr = "4";
													else if (dblCr <= 50) strCr = "5";
													else if (dblCr <= 60) strCr = "6";
													else if (dblCr <= 70) strCr = "7";
													else if (dblCr <= 80) strCr = "8";
													else if (dblCr <= 100) strCr = "9";
												}
											}
										}


										//damage codes
										if (this.m_dt.Rows[y]["fvs_dmg_ag1"] == System.DBNull.Value)
										{
										
											//cull board feet is the priority damage code
											if (this.m_dt.Rows[y]["cullbf"] != System.DBNull.Value)
											{
												dblCullBf = Convert.ToDouble(this.m_dt.Rows[y]["cullbf"]);
												if (dblCullBf > 0)
												{
													dblCullBf = Math.Round(dblCullBf,0);
													if (dblCullBf >= 100)
													{
														strDmgAg1 = "25";
														strDmgSv1 = "99";
													}
													else
													{
														strDmgAg1 = "25";
														strDmgSv1 = Convert.ToString(dblCullBf).Trim();
													}
												}
											}

											//dwarf mistletoe
											if (this.m_dt.Rows[y]["mist_cl_cd"] != System.DBNull.Value)
											{
												strMistClCd = Convert.ToString(this.m_dt.Rows[y]["mist_cl_cd"]).Trim();
												if (strMistClCd.Length > 0 && strMistClCd != "0")
												{
													/********************************************************
													 **attempt to convert FIA IDB dwarf mistletoe code 30
													 **to FVS dwarf mistletoe code that is based on species.
													 ********************************************************/
													
													if (intSpcRec != -1)
													{
														switch (strSpCd[intSpcRec,FVS_SPECIES].Trim())
														{
															case "LP":
																if (strDmgAg1=="0")
																{
																	strDmgAg1="31";
																	strDmgSv1=strMistClCd.Trim();
																}
																else
																{
																	strDmgAg2="31";
																	strDmgSv2=strMistClCd.Trim();
																}
																break;
															case "WL":
																if (strDmgAg1=="0")
																{
																	strDmgAg1="32";
																	strDmgSv1=strMistClCd.Trim();
																}
																else
																{
																	strDmgAg2="32";
																	strDmgSv2=strMistClCd.Trim();
																}
																break;
															case "DF":
																if (strDmgAg1=="0")
																{
																	strDmgAg1="33";
																	strDmgSv1=strMistClCd.Trim();
																}
																else
																{
																	strDmgAg2="33";
																	strDmgSv2=strMistClCd.Trim();
																}
																break;
															case "PP":
																if (strDmgAg1=="0")
																{
																	strDmgAg1="34";
																	strDmgSv1=strMistClCd.Trim();
																}
																else
																{
																	strDmgAg2="34";
																	strDmgSv2=strMistClCd.Trim();
																}
																break;
															default:
																if (strDmgAg1=="0")
																{
																	strDmgAg1="30";
																	strDmgSv1=strMistClCd.Trim();
																}
																else
																{
																	strDmgAg2="30";
																	strDmgSv2=strMistClCd.Trim();
																}

																break;

														}
													}
													else
													{
														if (strDmgAg1=="0")
														{
															strDmgAg1="30";
															strDmgSv1=strMistClCd.Trim();
														}
														else
														{
															strDmgAg2="30";
															strDmgSv2=strMistClCd.Trim();
														}
														
													}
												}
											}
											if (bBrokenTop==true)
											{
												if (strDmgAg1=="0")
												{
													strDmgAg1="96";
												}
												else if (strDmgAg2=="0")
												{
													strDmgAg2="96";
												}
												else
												{
													strDmgAg3="96";
												}
											}
											
										}
										else
										{
											//fvs damage code fields used in tree table
											strDmgAg1 = Convert.ToString(this.m_dt.Rows[y]["fvs_dmg_ag1"]).Trim();
											if (this.m_dt.Rows[y]["fvs_dmg_sv1"] != System.DBNull.Value)
											{
												if (Convert.ToString(this.m_dt.Rows[y]["fvs_dmg_sv1"]).Trim().Length > 0)
												{
													strDmgSv1 = Convert.ToString(this.m_dt.Rows[y]["fvs_dmg_sv1"]).Trim();
												}
											}
											if (this.m_dt.Rows[y]["fvs_dmg_ag2"] != System.DBNull.Value)
											{
												if (Convert.ToString(this.m_dt.Rows[y]["fvs_dmg_ag2"]).Trim().Length > 0 &&
													Convert.ToInt32(this.m_dt.Rows[y]["fvs_dmg_ag2"]) != Convert.ToInt32(strDmgAg1))
												{
													strDmgAg2 = Convert.ToString(this.m_dt.Rows[y]["fvs_dmg_ag2"]).Trim();
													if (this.m_dt.Rows[y]["fvs_dmg_sv2"] != System.DBNull.Value)
													{
														if (Convert.ToString(this.m_dt.Rows[y]["fvs_dmg_sv2"]).Trim().Length > 0)
														{
															strDmgSv2 = Convert.ToString(this.m_dt.Rows[y]["fvs_dmg_sv2"]).Trim();
														}
													}

												}
											}
											if (this.m_dt.Rows[y]["fvs_dmg_ag3"] != System.DBNull.Value)
											{
												if (Convert.ToString(this.m_dt.Rows[y]["fvs_dmg_ag3"]).Trim().Length > 0 &&
													Convert.ToInt32(this.m_dt.Rows[y]["fvs_dmg_ag3"]) != Convert.ToInt32(strDmgAg2) &&
													Convert.ToInt32(this.m_dt.Rows[y]["fvs_dmg_ag3"]) != Convert.ToInt32(strDmgAg1))
												{
													strDmgAg3 = Convert.ToString(this.m_dt.Rows[y]["fvs_dmg_ag3"]).Trim();
													if (this.m_dt.Rows[y]["fvs_dmg_sv3"] != System.DBNull.Value)
													{
														if (Convert.ToString(this.m_dt.Rows[y]["fvs_dmg_sv3"]).Trim().Length > 0)
														{
															strDmgSv3 = Convert.ToString(this.m_dt.Rows[y]["fvs_dmg_sv3"]).Trim();
														}
													}

												}
											}

										}
                                         
										if (strDmgAg1.Trim() == "25")
										{
											strValueClass = "3";
										}
										else if (Convert.ToInt32(strDmgSv1) > 0)
										{
											strValueClass = "2";
										}
										else
										{
											strValueClass="1";
										}


									}
								}
							}
							if (bSkip==false)
							{
                                if (intFvsTreeIdRecordCount > 0)
                                {
                                    //
                                    //PROCESS FVS_TREE_ID
                                    //
                                    if (FVS_TREE_ID.intPlotId == -1)
                                    {
                                        //check to see if any records in the fvs_tree_id_work_table
                                        if (intFvsTreeIdRecordCount > 0)
                                        {
                                            //the current tree does not have an assigned fvs_tree_id
                                            //so lets determine if the biosum_cond_id exists in the 
                                            //fvs_tree_id work table
                                            intPlotCnt = (int)m_ado.getSingleDoubleValueFromSQLQuery(m_ado.m_OleDbConnection, "SELECT COUNT(*) AS tempcount FROM fvs_tree_id_work_table WHERE biosum_cond_id='" + strRecordId.Trim() + "'", "temp");
                                            if (intPlotCnt > 0)
                                            {
                                                //biosum_cond_id is found so lets get its plot id and last tree id
                                                strPlotId = (string)m_ado.getSingleStringValueFromSQLQuery(
                                                    m_ado.m_OleDbConnection,
                                                    "SELECT DISTINCT plotid " +
                                                    "FROM fvs_tree_id_work_table " +
                                                    "WHERE biosum_cond_id='" + strRecordId.Trim() + "'", "temp");
                                                if (x == MORT_TREE)
                                                {
                                                    intTreeCnt = GetMaximumDeadTreeId(strRecordId) + 1;
                                                    if (intTreeCnt > 999)
                                                    {
                                                        //for dead trees valid treeid is 500 to 999.
                                                        //attempt to find an unused treeid between
                                                        //500 and 999
                                                        intTreeCnt = this.GetAvailableDeadTreeId(strRecordId);
                                                        if (intTreeCnt > 999)
                                                        {
                                                            m_intError = -1;
                                                            MessageBox.Show("Maximum tree id's have been reached for biosum_cond_id " + strRecordId, "FIA Biosum");
                                                            break;
                                                        }

                                                    }

                                                }
                                                else
                                                {
                                                    intTreeCnt = GetMaximumLiveTreeId(strRecordId) + 1;
                                                    if (intTreeCnt > 499)
                                                    {
                                                        //for live trees valid treeid is 001 to 499.
                                                        //attempt to find an unused treeid between
                                                        //001 and 499
                                                        intTreeCnt = this.GetAvailableLiveTreeId(strRecordId);
                                                        if (intTreeCnt > 499)
                                                        {
                                                            m_intError = -1;
                                                            MessageBox.Show("Maximum tree id's have been reached for biosum_cond_id " + strRecordId, "FIA Biosum");
                                                            break;
                                                        }

                                                    }
                                                }
                                                strTreeId = Convert.ToString(intTreeCnt).Trim();
                                                strTreeId = strTreeId.PadLeft(3, '0');


                                            }
                                            else
                                            {
                                                //okay, this is a new plot so 
                                                //lets get the last plot id and assign the 
                                                //first tree
                                                intPlotCnt = GetMaximumPlotId() + 1;
                                                if (intPlotCnt > 9999)
                                                {
                                                    intPlotCnt = GetAvailablePlotId();
                                                    if (intPlotCnt > 9999)
                                                    {
                                                        m_intError = -1;
                                                        MessageBox.Show("Maximum plot id's have been reached.", "FIA Biosum");
                                                        break;
                                                    }
                                                }
                                                strPlotId = Convert.ToString(intPlotCnt).Trim();
                                                strPlotId = strPlotId.PadLeft(4, '0');
                                                if (x == MORT_TREE)
                                                {
                                                    strTreeId = "500";
                                                }
                                                else
                                                {
                                                    strTreeId = "001";
                                                }
                                            }
                                        }
                                        else
                                        {
                                            strPlotId = "0001";
                                            if (x == MORT_TREE)
                                            {
                                                strTreeId = "500";
                                            }
                                            else
                                            {
                                                strTreeId = "001";
                                            }

                                        }
                                    }
                                    else
                                    {
                                        strPlotId = FVS_TREE_ID.strPlotId_formatted;
                                        strTreeId = FVS_TREE_ID.strTreeId_formatted;

                                    }
                                }

                                if (m_intError == 0)
                                {
                                    strDbh = strDbh.PadLeft(4, ' ');
                                    strHt = strHt.PadLeft(3, ' ');
                                    strActualHt = strActualHt.PadLeft(3, ' ');
                                    strDmgAg1 = strDmgAg1.PadLeft(2, ' ');
                                    strDmgSv1 = strDmgSv1.PadLeft(2, ' ');
                                    strDmgAg2 = strDmgAg2.PadLeft(2, ' ');
                                    strDmgSv2 = strDmgSv2.PadLeft(2, ' ');
                                    strDmgAg3 = strDmgAg3.PadLeft(2, ' ');
                                    strDmgSv3 = strDmgSv3.PadLeft(2, ' ');


                                    p_sw.WriteLine("{0,-4}{1,-3}{2,-6}{3,-1}{4,-3}{5,-4}{6,-3}{7,-3}{8,-3}{9,-4}{10,-1}{11,-2}{12,-2}{13,-2}{14,-2}{15,-2}{16,-2}{17,-1}{18,-1}",
                                        strPlotId, strTreeId, strTpa, strHistCd,
                                        strFvsCvtSpCd, strDbh, strDbhInc10yr, strHt, strActualHt, "    ", strCr,
                                        strDmgAg1, strDmgSv1, strDmgAg2, strDmgSv2, strDmgAg3, strDmgSv3, strValueClass, "0");

                                    strFVSTreeId = this.m_strVariant.Trim() +
                                        strPlotId.Trim() +
                                        strTreeId.Trim();
                                    if (FVS_TREE_ID.intPlotId == -1)
                                    {
                                        //
                                        //Save the newly assigned FVS_TREE_ID 
                                        //
                                        if (strCN.Trim().Length > 0)
                                        {
                                            this.m_ado.m_strSQL = "INSERT INTO " +
                                                this.m_strFVSTreeIdFIAWorkTable + " " +
                                                "(biosum_cond_id,cn,fvs_tree_id,fvs_filename) VALUES (" +
                                                "'" + strCurRecordId.Trim() + "'," +
                                                "'" + strCN.Trim() + "'," +
                                                "'" + strFVSTreeId.Trim() + "'," +
                                                "'" + strFVSFile.Trim() + "');";
                                        }
                                        else
                                        {
                                            this.m_ado.m_strSQL = "INSERT INTO " +
                                                this.m_strFVSTreeIdIDBWorkTable + " " +
                                                "(biosum_cond_id,idb_alltree_id,fvs_tree_id,fvs_filename) VALUES (" +
                                                "'" + strCurRecordId.Trim() + "'," +
                                                strIDBAlltreeId.Trim() + "," +
                                                "'" + strFVSTreeId.Trim() + "'," +
                                                "'" + strFVSFile.Trim() + "');";
                                        }
                                        this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection, this.m_ado.m_strSQL);
                                        if (intFvsTreeIdRecordCount > 0)
                                        {
                                            m_ado.m_strSQL = "INSERT INTO fvs_tree_id_work_table " +
                                                           "(biosum_cond_id,fvs_tree_id,fvs_variant," +
                                                            "plotid,treeid,plottreeid) VALUES " +
                                                           "('" + strCurRecordId.Trim() + "'," +
                                                            "'" + strFVSTreeId.Trim() + "'," +
                                                            "'" + m_strVariant + "'," +
                                                            "'" + strPlotId + "'," +
                                                            "'" + strTreeId + "'," +
                                                            "'" + strPlotId + strTreeId + "')";
                                            this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection, this.m_ado.m_strSQL);
                                        }
                                        

                                    }
                                }
							}
                            if (m_intError < 0) break;
						
						}
					}
				}

                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text", "Finishing Up With FVS Variant " + this.m_strVariant.Trim() + "...Stand By");
                if (m_intError == 0)
                {
                    if (Convert.ToInt32(this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection, "select count(*) from " + this.m_strFVSTreeIdIDBWorkTable, this.m_strFVSTreeIdIDBWorkTable)) > 0)
                    {
                        this.m_ado.m_strSQL = "UPDATE " + this.m_strTreeTable + " t " +
                            "INNER JOIN " + this.m_strFVSTreeIdIDBWorkTable + " w " +
                            "ON t.idb_alltree_id = w.idb_alltree_id " +
                            "SET t.fvs_tree_id=w.fvs_tree_id;";
                        this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection, this.m_ado.m_strSQL);
                        this.m_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " +
                                              "INNER JOIN " + this.m_strFVSTreeIdIDBWorkTable + " w " +
                                              "ON c.biosum_cond_id = w.biosum_cond_id " +
                                              "SET c.fvs_filename = mid(w.fvs_filename,1,12);";
                        this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection, this.m_ado.m_strSQL);

                    }
                    if (Convert.ToInt32(this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection, "select count(*) from " + this.m_strFVSTreeIdFIAWorkTable, this.m_strFVSTreeIdFIAWorkTable)) > 0)
                    {
                        this.m_ado.m_strSQL = "UPDATE " + this.m_strTreeTable + " t " +
                            "INNER JOIN " + this.m_strFVSTreeIdFIAWorkTable + " w " +
                            "ON t.cn = w.cn " +
                            "SET t.fvs_tree_id=w.fvs_tree_id;";
                        this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection, this.m_ado.m_strSQL);
                        this.m_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " +
                            "INNER JOIN " + this.m_strFVSTreeIdFIAWorkTable + " w " +
                            "ON c.biosum_cond_id = w.biosum_cond_id " +
                            "SET c.fvs_filename = mid(w.fvs_filename,1,12);";
                        this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection, this.m_ado.m_strSQL);
                    }
                }
				
																	

			}
			catch (Exception e)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - fvs_input:CreateFVS  \n" + 
					"Err Msg - " + e.Message.ToString().Trim(),
					"FVS Input",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}
			if (this.m_ado.m_OleDbConnection != null)
				this.m_ado.m_OleDbConnection.Close();
			p_sw.Close();
			p_fs.Close();
			p_sw=null;
			p_fs=null;
		}
        
		private void CheckDir()
		{
			try
			{
		
				if (!System.IO.Directory.Exists(this.m_strInDir))
					System.IO.Directory.CreateDirectory(this.m_strInDir);
			}
			catch (Exception e)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - fvs_input:CheckDir  \n" + 
					"Err Msg - " + e.Message.ToString().Trim(),
					"FVS Input",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}
		}
		private string IDB_getInvYear(string p_strInvId)
		{
			switch (p_strInvId.Trim())
			{
				case "0010" :
					return "1993";
				case "0011" :
					return "1993";
				case "0012" :
				    return "1997";
				case "0013" :
					return "1998";
				case "0014" :
					return "1995";
				case "0015" :
					return "1991";
				case "0016" :
					return "1991";
				case "0017" :
					return "1988";
				case "0018" :
					return "   @";
				case "0019" :
					return "   @";
			}
			return "   @";
		}
		private int IDB_getStandAge(int intStdAge)
		{
			switch (intStdAge)
			{
				case 1:	return 5;
				case 2:	return 15;
				case 3:	return 25;
				case 4:	return 35;
				case 5:	return 45;
				case 6:	return 55;
				case 7:	return 65;
				case 8:	return 75;
				case 9:	return 85;
				case 10: return 95;
				case 11: return 105;
				case 12: return 115;
				case 13: return 125;
				case 14: return 135;
				case 15: return 145;
				case 16: return 155;
				case 17: return 165;
				case 18: return 175;
				case 19: return 185;
				case 20: return 195;
				case 21: return 250;
				case 22: return  350;
			}
			return 0;
		}
		private string IDB_getStateCd(string p_strStateCd)
		{
			switch (p_strStateCd)
			{
				case "41":	return "OR"; 
				case "53":  return "WA";
				case "6":   return "CA";
				case "1":   return "AL";
				case "2":   return "AK";
				case "4":   return "AZ";
				case "5":   return "AR";
				case "8":   return "CO";
				case "9":   return "CT";
				case "10":  return "DE";
				case "11":  return "DC";
				case "12":  return "FL";
				case "13":  return "GA";
				case "15":  return "HI";
				case "16":  return "ID";
				case "17":  return "IL";
				case "18":  return "IN";
				case "19":  return "IA";
				case "20":  return "KS";
				case "21":  return "KY";
				case "22":  return "LA";
				case "23":  return "ME";
				case "24":  return "MD";
				case "25":  return "MA";
				case "26":  return "MI";
				case "27":  return "MN";
				case "28":  return "MS";
				case "29":  return "MO";
				case "30":  return "MT";
				case "31":  return "NE";
				case "32":  return "NV";
				case "33":  return "NH";
				case "34":  return "NJ";
				case "35":  return "NM";
				case "36":  return "NY";
				case "37":  return "NC";
				case "38":  return "ND";
				case "39":  return "OH";
				case "40":  return "OK";
				case "42":  return "PA";
				case "44":  return "RI";
				case "45":  return "SC";
				case "46":  return "SD";
				case "47":  return "TN";
				case "48":  return "TX";
				case "49":  return "UT";
				case "50":  return "VT";
				case "51":  return "VA";
				case "54":  return "WV";
				case "55":  return "WI";
				case "56":  return "WY";
				case "72":  return "PR";
				case "78":  return "VI";
			}
			return "";
		}
		/// <summary>
		/// delete loc, slf, and fvs files from the variant directory
		/// </summary>
		private void DeleteFiles()
		{

			string strCurrDir = System.IO.Directory.GetCurrentDirectory();
			System.IO.Directory.SetCurrentDirectory(this.m_strInDir);
			string[] strFiles = System.IO.Directory.GetFiles(this.m_strInDir,"*.fvs");
			foreach(string strFile in strFiles)
			{
				System.IO.File.Delete(strFile.Trim());
			}
			System.IO.File.Delete(this.m_strVariant.Trim() + ".loc");
			System.IO.File.Delete(this.m_strVariant.Trim() + ".slf");
			System.IO.Directory.SetCurrentDirectory(strCurrDir);

		}
		private void InitializeFields()
		{
			this.m_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
				                   " INNER JOIN " + this.m_strPlotTable + " p " + 
				                   " ON c.biosum_plot_id = p.biosum_plot_id " + 
				                   " SET c.fvs_filename = NULL " + 
				                   " WHERE TRIM(p.fvs_variant)='" + this.m_strVariant.Trim() + "';";
			this.m_ado.SqlNonQuery(this.m_strConn,this.m_ado.m_strSQL);
			if (this.m_ado.m_intError != 0) this.m_intError = -1;
		}
		/// <summary>
		/// full directory path and file name to the fvsin mdb file
		/// </summary>
		/// <returns>string name of the Input MDB File</returns>
		public string strFVSInMDBFile
		{
		    set
			{
				this.m_strFVSInMDBFile = value;
			}
			get
			{
				return this.m_strFVSInMDBFile;
			}


		}
        /// <summary>
        /// Get the maximum dead tree id for a biosum_cond_id
        /// </summary>
        /// <param name="p_strBiosum_Cond_Id"></param>
        /// <returns></returns>
        private int GetMaximumDeadTreeId(string p_strBiosum_Cond_Id)
        {
            int intValue = (int)this.m_ado.getSingleDoubleValueFromSQLQuery(
               m_ado.m_OleDbConnection,
               "SELECT MAX(VAL(treeid)) FROM fvs_tree_id_work_table WHERE biosum_cond_id='" +
               p_strBiosum_Cond_Id.Trim() + "' AND VAL(treeid) > 499",
               "fvs_tree_id_work_table");
            if (intValue == null) intValue = -1;
            return intValue;
        }
        /// <summary>
        /// Get the maximum live tree id for a biosum_cond_id
        /// </summary>
        /// <param name="p_strBiosum_Cond_Id"></param>
        /// <returns></returns>
        private int GetMaximumLiveTreeId(string p_strBiosum_Cond_Id)
        {
            int intValue = (int)this.m_ado.getSingleDoubleValueFromSQLQuery(
               m_ado.m_OleDbConnection,
               "SELECT MAX(VAL(treeid)) FROM fvs_tree_id_work_table WHERE biosum_cond_id='" + 
               p_strBiosum_Cond_Id.Trim() + "' AND VAL(treeid) < 500",
               "fvs_tree_id_work_table");
            if (intValue == null) intValue = -1;
            return intValue;
        }
        /// <summary>
        /// Get the maximum assigned plot id in order to assign the next new plot id
        /// </summary>
        /// <returns></returns>
        private int GetMaximumPlotId()
        {
            int intValue = (int)this.m_ado.getSingleDoubleValueFromSQLQuery(
                m_ado.m_OleDbConnection, 
                "SELECT MAX(VAL(plotid))  as max_plotid FROM fvs_tree_id_work_table",
                "fvs_tree_id_work_table");
            if (intValue == null) intValue = -1;
            return intValue;
        }
        /// <summary>
        /// find the 1st unused plot id
        /// </summary>
        /// <returns></returns>
        private int GetAvailablePlotId()
        {
            int intValue = -1;
            string strPlotId = "";
            int x;
            for (x = 1; x <= 9999; x++)
            {
                strPlotId = Convert.ToString(x).PadLeft(4,'0');
                intValue = (int)m_ado.getSingleDoubleValueFromSQLQuery(m_ado.m_OleDbConnection,
                    "SELECT COUNT(*) as plotid_count FROM fvs_tree_id_work_table " + 
                    "WHERE plotid='" + strPlotId.Trim() + "'","temp");
                if (intValue == 0)
                {
                    break;
                }
            }
            intValue = x;
            return intValue;
        }

        /// <summary>
        /// find the 1st unused live tree id
        /// </summary>
        /// <returns></returns>
        private int GetAvailableLiveTreeId(string p_strBiosum_Cond_Id)
        {
            int intValue = -1;
            string strTreeId = "";
            int x;
            for (x = 1; x <= 499; x++)
            {
                strTreeId = Convert.ToString(x).PadLeft(3, '0');
                intValue = (int)m_ado.getSingleDoubleValueFromSQLQuery(m_ado.m_OleDbConnection,
                    "SELECT COUNT(*) as treeid_count FROM fvs_tree_id_work_table " +
                    "WHERE biosum_cond_id='" + p_strBiosum_Cond_Id + "' AND " + 
                          "treeid='" + strTreeId.Trim() + "'", "temp");
                if (intValue == 0)
                {
                    break;
                }
            }
            intValue = x;
            return intValue;
        }
        /// <summary>
        /// find the 1st unused dead tree id
        /// </summary>
        /// <returns></returns>
        private int GetAvailableDeadTreeId(string p_strBiosum_Cond_Id)
        {
            int intValue = -1;
            string strTreeId = "";
            int x;
            for (x = 500; x <= 999; x++)
            {
                strTreeId = Convert.ToString(x).PadLeft(3, '0');
                intValue = (int)m_ado.getSingleDoubleValueFromSQLQuery(m_ado.m_OleDbConnection,
                    "SELECT COUNT(*) as treeid_count FROM fvs_tree_id_work_table " +
                    "WHERE biosum_cond_id='" + p_strBiosum_Cond_Id + "' AND " +
                          "treeid='" + strTreeId.Trim() + "'", "temp");
                if (intValue == 0)
                {
                    break;
                }
            }
            intValue = x;
            return intValue;
        }

        /// <summary>
        /// Site index functions beginning with "z" were programmed by Tara: these have been checked against the publications.
        /// Site index functions beginning with "q" were taken from Bruce Hiserote's Visual Basic program 4/2002: these should be fine.
        ///Other site index functions taken from FSVEG group - Kurt Campbell: these have not been checked for errors.
        ///Any changes from the sources are noted in the comments of each function.
        ///Other Site index equations come from the publication
        ///"Site Index Equations and Mean Annual Increment Equations
        ///for Pacific Northwest Research Station Forest Inventory and 
        ///Analysis Inventories, 1985-2001" 
        ///Authors: Erica Hanson, David Azuma, and Bruce Hiserote
        ///Research Note: PNW-RN533 December 2002
        /// </summary>
        private class site_index
		{
			double _dblCCTreeBasalAreaPerAcre;
			double _dblCCAvgDia;
			string _strStateCd="";
			string _strCountyCd="";
			string _strPlot="";
			string _strCondId="";
			string _strBiosumPlotId="";
			ado_data_access _oAdo;
			string _strPlotTable="";
			string _strTreeTable="";
			string _strCondTable="";
			string _strSiteTreeTable="";
			string _strTreeSpeciesTable="";
			string _strFVSTreeSpcTable="";
			string _strFVSVariant="";
			string _strSiteIndexSpecies="";
			string _strSiteIndex="";
			string _strSiteIndexSpeciesAlphaCode="";
            string _strCCHabitatTypeCd;
            IDictionary<String, String> _dictSiteIdxEq;
            string _strDebugFile;
			
			bool _bProcess=true;

			public site_index()
			{
			}
			
			public string StateCd
			{
				get {return _strStateCd;}
				set {_strStateCd=value;}
			}
			public string CountyCd
			{
				get {return _strCountyCd;}
				set {_strCountyCd=value;}
			}
			public string Plot
			{
				get {return _strPlot;}
				set {_strPlot=value;}
			}
			public string CondId
			{
				get {return _strCondId;}
				set {_strCondId=value;}
			}
			public string BiosumPlotId
			{
				get {return _strBiosumPlotId;}
				set {_strBiosumPlotId=value;}
			}
			public string FVSVariant
			{
				get {return _strFVSVariant;}
				set {_strFVSVariant=value;}
			}
			public string PlotTable
			{
				get {return _strPlotTable;}
				set {_strPlotTable=value;}
			}
			public string CondTable
			{
				get {return _strCondTable;}
				set {_strCondTable=value;}
			}
			public string TreeTable
			{
				get {return _strTreeTable;}
				set {_strTreeTable=value;}
			}
			public string SiteTreeTable
			{
				get {return _strSiteTreeTable;}
				set {_strSiteTreeTable=value;}
			}
			public string TreeSpeciesTable
			{
				get {return _strTreeSpeciesTable;}
				set {_strTreeSpeciesTable=value;}
			}
			public string FVSTreeSpeciesTable
			{
				get {return _strFVSTreeSpcTable;}
				set {_strFVSTreeSpcTable=value;}
			}
			public ado_data_access ado_data_access
			{
				set {this._oAdo=value;}
				get {return _oAdo;}
			}
			public bool Process
			{
				set {_bProcess=value;}
				get {return _bProcess;}
			}
			public string SiteIndexSpecies
			{
				get {return _strSiteIndexSpecies;}
				set {_strSiteIndexSpecies=value;}
			}
			public string SiteIndex
			{
				get {return _strSiteIndex;}
				set {_strSiteIndex=value;}
			}
			public string SiteIndexSpeciesAlphaCode
			{
				get {return _strSiteIndexSpeciesAlphaCode;}
				set {_strSiteIndexSpeciesAlphaCode=value;}
			}
			public double ConditionClassBasalAreaPerAcre
			{
				get {return _dblCCTreeBasalAreaPerAcre;}
				set {_dblCCTreeBasalAreaPerAcre=value;}
			}
			public double ConditionClassAverageDia
			{
				get {return _dblCCAvgDia;}
				set {_dblCCAvgDia=value;}
			}
            public string ConditionClassHabitatTypeCd
            {
                get { return _strCCHabitatTypeCd; }
                set { _strCCHabitatTypeCd = value; }
            }
            public IDictionary<String, String> SiteIndexEquations
            {
                set { _dictSiteIdxEq = value; }
            }
            public string DebugFile
            {
                set { _strDebugFile = value; }
            }
 
			public void getSiteIndex(System.Data.DataRow p_oRow)
			{
				//biosum plot id
				if (p_oRow["biosum_plot_id"] == System.DBNull.Value)
					this.BiosumPlotId="";
				else this.BiosumPlotId=p_oRow["biosum_plot_id"].ToString().Trim();
				//statecd
				if (p_oRow["statecd"] == System.DBNull.Value)
					this.StateCd="";
				else this.StateCd=Convert.ToString(p_oRow["statecd"]).Trim();
				//countycd
				if (p_oRow["countycd"] == System.DBNull.Value)
					this.CountyCd="";
				else this.CountyCd=Convert.ToString(p_oRow["countycd"]).Trim();
				//plot
				if (p_oRow["plot"] == System.DBNull.Value)
					this.Plot="";
				else this.Plot=Convert.ToString(p_oRow["plot"]).Trim();
				//fvs variant
				if (p_oRow["fvs_variant"] == System.DBNull.Value)
					this.FVSVariant="";
				else this.FVSVariant=Convert.ToString(p_oRow["fvs_variant"]).Trim();
				//cond id
				if (p_oRow["condid"] == System.DBNull.Value)
					this.CondId="";
				else this.CondId=Convert.ToString(p_oRow["condid"]).Trim();
                //habitat type code
                if (p_oRow["habtypcd1"] == System.DBNull.Value)
                    this.ConditionClassHabitatTypeCd = "";
                else this.ConditionClassHabitatTypeCd = Convert.ToString(p_oRow["habtypcd1"]).Trim();

				//15-JUN-2015: We now calculate this in a function rather than populate from COND table so we can control the parameters
                //tree basal area per acre on the condition
                //if (p_oRow["ba_ft2_ac"] != System.DBNull.Value)
                //    this.ConditionClassBasalAreaPerAcre=Convert.ToDouble(p_oRow["ba_ft2_ac"]);

				getSiteIndex();
			}
			private void getSiteIndex()
			{
				int x,y;
				int intCount;

				//MAX variables hold the values of the selected site index
				int intSICountMax;
				double dblSIAvgMax;
				int intSISpeciesMax;
				int intCurSIFVSSpecies;
				int intCurFIASpecies;
				int intCurHtFt;
				int intCurAgeDia;
				int intCondId;
				bool bFound;
                int intSiTree;

				//These arrays contain the values of all the site index trees on the plot
				int[] intSIFVSSpecies;
				int[] intSICount;
				double[] dblSISum;
				double[] dblSIAvg;

				this.SiteIndex="@";
				this.SiteIndexSpecies="@";
				this.SiteIndexSpeciesAlphaCode="@";

				double dblSiteIndex=0;


				//calculate site index for OR, WA, CA, ID, and MT
                if (StateCd=="41" || StateCd=="6" ||
					StateCd=="53" || StateCd=="16" ||
                    StateCd=="30")
				{
				}
				else return;

				//get all the site index trees that are applied to the current plot+condition
				_oAdo.m_strSQL = "SELECT s.biosum_plot_id," + 
					"s.condid," + 
					"s.tree," + 
					"s.spcd," + 
					"s.dia," + 
					"s.ht," + 
					"s.agedia," + 
					"s.subp," + 
					"s.method," + 
					"s.validcd, " +
                    "s.sitree " +
					"FROM " + this.SiteTreeTable + " s " + 
					"WHERE s.biosum_plot_id = '" + this.BiosumPlotId + "' " +
                    "AND s.condid = " + this.CondId +
                    "AND s.validcd <> 0";
				x=Convert.ToInt32(_oAdo.getRecordCount(_oAdo.m_OleDbConnection,"SELECT COUNT(*) FROM (" + _oAdo.m_strSQL + ")","cond"));
				if (x> 0)
				{
					_oAdo.AddSQLQueryToDataSet(_oAdo.m_OleDbConnection,ref _oAdo.m_OleDbDataAdapter,ref _oAdo.m_DataSet,_oAdo.m_strSQL,"GetSiteIndex");
					if (_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows.Count > 0)
					{
						intSIFVSSpecies = new int[x];
						intSICount = new int[x];
						dblSISum = new double[x];
						intCount=0;
						intSICountMax=0;
						for (y=0;y<=_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows.Count-1;y++)
						{
							intCurFIASpecies= Convert.ToInt32(_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows[y]["spcd"]);
							intCurSIFVSSpecies = 0;
							intCurAgeDia = Convert.ToInt32(_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows[y]["agedia"]);
							intCurHtFt = Convert.ToInt32(_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows[y]["ht"]);
							intCondId = Convert.ToInt32(_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows[y]["condid"]);
                            intSiTree = Convert.ToInt32(_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows[y]["sitree"]);

							//***************************************************
							//**if no age then bypass site index tree
							//**************************************************
							if (intCurAgeDia > 0)
							{
								LoadSiteIndexValues(intCondId,intCurFIASpecies, 
									intCurAgeDia, 
									intCurHtFt,
									ref intCurSIFVSSpecies,
									ref dblSiteIndex,
                                    intSiTree);
                                //*************************************************
                                //**if the site index = 0, write it to the log, we want to know
                                //**how often this occurs
                                //*************************************************
                                if (dblSiteIndex == 0)
                                {
                                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    {
                                        string logEntry = "//variant: " + this.FVSVariant +
                                                          " plot id: " + this.BiosumPlotId +
                                                          " cond id: " + this.CondId +
                                                          " spec cd: " + intCurSIFVSSpecies + "\r\n";
                                        frmMain.g_oUtils.WriteText(_strDebugFile, "\r\n//\r\n");
                                        frmMain.g_oUtils.WriteText(_strDebugFile, "//Site_Index_getSiteIndex\r\n");
                                        frmMain.g_oUtils.WriteText(_strDebugFile, "//Site index equation returned 0    \r\n");
                                        frmMain.g_oUtils.WriteText(_strDebugFile, logEntry);
                                        frmMain.g_oUtils.WriteText(_strDebugFile, "//\r\n");
                                    }
                                }
								//*************************************************
								//**lets find the current SI species in the array
								//*************************************************
								if (intCount==0)
								{
									intCount=intCount+1;
									intSIFVSSpecies[intCount-1]=intCurSIFVSSpecies;
									intSICount[intCount-1]=intSICount[intCount-1] +  1;
									dblSISum[intCount-1] = dblSISum[intCount-1] + dblSiteIndex;
									
									
								}
								else if (intSIFVSSpecies[intCount-1]==intCurSIFVSSpecies)
								{
									intSICount[intCount-1] = intSICount[intCount-1] + 1;
									dblSISum[intCount-1] = dblSISum[intCount-1] + dblSiteIndex;
								}
								else
								{
									bFound=false;
									for (x=0;x<=intCount-1;x++)
									{
										if (intSIFVSSpecies[x] == intCurSIFVSSpecies)
										{
											bFound=true;
											break;
										}
									}
									if (bFound)
									{
										intSICount[x] = intSICount[x] + 1;
										dblSISum[x] = dblSISum[x] + dblSiteIndex;
									}
									else
									{
										intCount=intCount+1;
										intSIFVSSpecies[intCount-1]=intCurSIFVSSpecies;
										intSICount[intCount-1]=intSICount[intCount-1] +  1;
										dblSISum[intCount-1] = dblSISum[intCount-1] + dblSiteIndex;
									}
								}
							}
						}
						//***************************************************************
						//**get the most frequently occuring site index species
						//***************************************************************
						intSICountMax = 0;
						dblSIAvgMax = 0;
						intSISpeciesMax=0;
						dblSIAvg = new double[intCount];
						for (x=0;x<=intCount-1;x++)
						{
							dblSIAvg[x] = dblSISum[x] / intSICount[x];
							if (intSICount[x] > intSICountMax)
							{
								intSICountMax = intSICount[x];
								dblSIAvgMax = dblSIAvg[x];
								intSISpeciesMax = intSIFVSSpecies[x];
							}
							else if (intSICount[x] == intSICountMax)
							{
								if (dblSIAvg[x] > dblSIAvgMax)
								{
									dblSIAvgMax = dblSIAvg[x];
									intSISpeciesMax = intSIFVSSpecies[x];
								}
							}
						}
						if (dblSIAvgMax<=0 && intSISpeciesMax > 0 && intSISpeciesMax != 999)
						{
							MessageBox.Show("Warning: Site index tree species " + Convert.ToString(intSISpeciesMax) + " has an invalid  site index value of " +  Convert.ToString(Math.Round(dblSIAvgMax,6)).Trim() + ". Both the SI species and SI height will be given a value of @");
							this.SiteIndexSpecies="@";
							this.SiteIndex="@";

						}
						else
						{
							this.SiteIndexSpecies = intSISpeciesMax.ToString().Trim();
							this.SiteIndex = Convert.ToString(Math.Round(dblSIAvgMax,0)).Trim();
						}
						if (this.SiteIndexSpecies != "@" && this.SiteIndexSpecies.Trim().Length > 0)
							GetSiteIndexSpeciesAlphaCode();
					}
					_oAdo.m_DataSet.Tables.Remove("GetSiteIndex");

				}
			}
			private void GetSiteIndexSpeciesAlphaCode()
			{
				_oAdo.m_strSQL = "SELECT fvs_species FROM " + this.FVSTreeSpeciesTable + " f "  + 
					             "WHERE fvs_variant = '" + this.FVSVariant + "' AND " +
					                   "spcd = " + this.SiteIndexSpecies ;

				this.SiteIndexSpeciesAlphaCode=_oAdo.getSingleStringValueFromSQLQuery(this._oAdo.m_OleDbConnection,_oAdo.m_strSQL,this.FVSTreeSpeciesTable);
				if (this.SiteIndexSpeciesAlphaCode.Trim().Length == 0)
					this.SiteIndexSpeciesAlphaCode="@";

				

			}

		
			private void LoadSiteIndexValues(int p_intSICondId,
				int p_intSISpCd,
				int p_intSIAgeDia,
				int p_intSIHtFt,
				ref int p_intSIFVSSpecies,
				ref double p_dblSiteIndex,
                int p_intSiTree)
			{
				
				
				//
				//Western Cascades variant
				//
				if (this.FVSVariant=="WC")
				{

					if (p_intSISpCd==11) //pacific silver fir
					{
						p_dblSiteIndex = ABAM2(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=11;
					}
					else if (p_intSISpCd==17 || p_intSISpCd==15)  //grand fir or white fir
					{
						p_dblSiteIndex = ABGR1(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=15;
					}
					else if (p_intSISpCd==19 || p_intSISpCd==93) //subalpine fir, englemann spruce
					{
						p_dblSiteIndex = PIEN3(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=93;
					}
					else if (p_intSISpCd==20 || p_intSISpCd==21)  //CA red fir, Shasta red fir
					{
						p_dblSiteIndex = zABMA2(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=20;
					}
					else if (p_intSISpCd==22) //noble fir
					{
						p_dblSiteIndex = ABPR1(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=22;
					}
					else if (p_intSISpCd==42 ||
						p_intSISpCd==119 || 
						p_intSISpCd==202 || 
						p_intSISpCd==242)  //Alaska cedar, western white pine, Douglas-fir, red cedar
					{
						p_dblSiteIndex = qPSME13(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=202;
					}
					else if (p_intSISpCd==108) //lodgepole
					{
						getAvgDbhAndBasalArea(p_intSICondId);
						p_dblSiteIndex = zPICO3(p_intSIAgeDia,
							p_intSIHtFt,
							this.ConditionClassBasalAreaPerAcre,
							this.ConditionClassAverageDia);
						p_intSIFVSSpecies=108;
					}
					else if (p_intSISpCd==119) //western white pine
					{
						p_dblSiteIndex = PIMO3(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies = 119;
					}
					else if (p_intSISpCd==122)  //ponderosa pine
					{
						p_dblSiteIndex = PIPO3(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=122;
					}
					else if (p_intSISpCd==263) //western hemlock
					{
						p_dblSiteIndex = TSHE1(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies= 263;
					}
					else if (p_intSISpCd==264) //mountain hemlock
					{
						p_dblSiteIndex = zTSME(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=264;
					}
					else if (p_intSISpCd==351) //red alder
					{
						p_dblSiteIndex = ALRU2(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=351;
					}
					else if (p_intSISpCd==72 || p_intSISpCd==73) //subalpine larch (larix lyallii)
					{
						p_dblSiteIndex= LAOC1_OR(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies= 73;
					}
					else if (p_intSISpCd==211 || 
						p_intSISpCd==312 ||
						p_intSISpCd==321 ||
						p_intSISpCd==475 ||
						p_intSISpCd==352 ||
						p_intSISpCd==375 ||
						p_intSISpCd==431 ||
						p_intSISpCd==746 ||
						p_intSISpCd==747 ||
						p_intSISpCd==815 ||
						p_intSISpCd==64  ||
						p_intSISpCd==101 ||
						p_intSISpCd==103 ||
						p_intSISpCd==231 ||
						p_intSISpCd==492 ||
						p_intSISpCd==500 ||
						p_intSISpCd==920)
					{
						//redwood, maple, maple, maple, white alder,
						//paper birch, golden chink, quaking aspen
						//black cottonwood, white oak, juniper, whitebark pine,
						//knobcone pine, pacific yew
						//pacific dogwood, hawthorne, bitter cherry, willow
						p_dblSiteIndex = qPSME13(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies = 202;
					}
					else
					{
						p_dblSiteIndex = 0;
						p_intSIFVSSpecies=999;
					}
					return;
				}
				//
				//Eastern Cascades variant
				//
				if (this.FVSVariant=="EC")
				{
					if (p_intSISpCd==119) //western white pine
					{
						p_dblSiteIndex = PIMO2(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=119;
					}
					else if (p_intSISpCd==73)  //western larch
					{
						p_dblSiteIndex= LAOC1_OR(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=73;
					}
					else if (p_intSISpCd==202) //Douglas-fir
					{
						p_dblSiteIndex=qPSME12(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=202;
					}
					else if (p_intSISpCd==11 || 
						p_intSISpCd==17 || 
						p_intSISpCd==15)	  //Pacific silver fir and grand fir
					{
						p_dblSiteIndex=ABGR1(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=15;
					}
					else if (p_intSISpCd==108) //lodgepole
					{
						getAvgDbhAndBasalArea(p_intSICondId);
						p_dblSiteIndex = zPICO3(p_intSIAgeDia,
							p_intSIHtFt,
							this.ConditionClassBasalAreaPerAcre,
							this.ConditionClassAverageDia);
						p_intSIFVSSpecies=108;
					}
					else if (p_intSISpCd==93) //englemann spruce
					{
						p_dblSiteIndex = PIEN3(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=93;
					}
					else if (p_intSISpCd==19 || p_intSISpCd==22) //subalpine fir,noble fir
					{
						p_dblSiteIndex = zABPR2(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=22;
					}
					else if (p_intSISpCd==122)  //ponderosa pine
					{
						p_dblSiteIndex = PIPO3(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=122;
					}
					else if (p_intSISpCd==264) //mountain hemlock
					{
						p_dblSiteIndex = zTSME(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=264;
					}
					
					else
					{
						p_dblSiteIndex = 0;
						p_intSIFVSSpecies=999;
					//	MessageBox.Show("No site index equation found for species " + p_intSISpCd.ToString() + " of variant " + this.FVSVariant);
					}

				}
				//
				//Pacific Northwest Coast variant
				//
				if (this.FVSVariant=="PN")
				{
					if (p_intSISpCd==11) //pacific silver fir
					{
						p_dblSiteIndex = ABAM2(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=11;
					}
					else if (p_intSISpCd==17 || p_intSISpCd==15)  //grand fir or white fir
					{
						p_dblSiteIndex = ABGR1(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=15;
					}
					else if (p_intSISpCd==19 || p_intSISpCd==93) //subalpine fir, englemann spruce
					{
						p_dblSiteIndex = PIEN3(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=93;
					}
					else if (p_intSISpCd==20 || p_intSISpCd==21)  //CA red fir, Shasta red fir
					{
						p_dblSiteIndex = zABMA2(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=20;
					}
					else if (p_intSISpCd==98) //Sitka spruce
					{
						p_dblSiteIndex = PISI1(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies= 98;
					}
					else if (p_intSISpCd==22) //noble fir
					{
						p_dblSiteIndex = ABPR1(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=22;
					}
					else if (p_intSISpCd==108) //lodgepole
					{
						getAvgDbhAndBasalArea(p_intSICondId);
						p_dblSiteIndex = zPICO3(p_intSIAgeDia,
							p_intSIHtFt,
							this.ConditionClassBasalAreaPerAcre,
							this.ConditionClassAverageDia);
						p_intSIFVSSpecies=108;
					}
					else if (p_intSISpCd==119) //western white pine
					{
						p_dblSiteIndex = PIMO3(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies = 119;
					}
					else if (p_intSISpCd==122)  //ponderosa pine
					{
						p_dblSiteIndex = PIPO3(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=122;
					}
					else if (p_intSISpCd==202) //Douglas-fir
					{
						p_dblSiteIndex = qPSME1(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=202;
					}
					else if (p_intSISpCd==263) //western hemlock
					{
						p_dblSiteIndex = TSHE1(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies= 263;
					}
					else if (p_intSISpCd==264) //mountain hemlock
					{
						p_dblSiteIndex = zTSME(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=264;
					}
					else if (p_intSISpCd==73 || p_intSISpCd==72)  //western larch
					{
						p_dblSiteIndex= LAOC1_OR(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=73;
					}
					else if (p_intSISpCd==42  ||
						p_intSISpCd==73  || 
						p_intSISpCd==211 ||
						p_intSISpCd==312 ||
						p_intSISpCd==321 ||
						p_intSISpCd==475 ||
						p_intSISpCd==352 ||
						p_intSISpCd==375 ||
						p_intSISpCd==431 ||
						p_intSISpCd==746 ||
						p_intSISpCd==747 ||
						p_intSISpCd==815 ||
						p_intSISpCd==64  ||
						p_intSISpCd==101 ||
						p_intSISpCd==103 ||
						p_intSISpCd==231 ||
						p_intSISpCd==492 ||
						p_intSISpCd==500 ||
						p_intSISpCd==768 ||
						p_intSISpCd==920)
					{	
					  //Alaska cedar,western larch, redwood, maple, maple, maple, white alder,
				      //paper birch, golden chink, quaking aspen
					  // black cottonwood, white oak, juniper, whitebark pine, knobcone pine, pacific yew
					  //pacific dogwood, hawthorne, bitter cherry, willow
						p_dblSiteIndex = qPSME13(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies = 202;

					}
					else
					{
						p_dblSiteIndex = 0;
						p_intSIFVSSpecies=999;
					}
				}
				//
				//Blue Mountains variant
				//
				if (this.FVSVariant=="BM")
				{
					if (p_intSISpCd==119) //western white pine
					{
						p_dblSiteIndex = PIMO2(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=119;
					}
					else if (p_intSISpCd==73)  //western larch
					{
						p_dblSiteIndex= LAOC1_OR(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=73;
					}
					else if (p_intSISpCd==202) //Douglas-fir
					{
						p_dblSiteIndex=qPSME12(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=202;
					}
                    else if (p_intSISpCd == 17 ||
                             p_intSISpCd == 15) //grand fir and white fir
					{
                        p_dblSiteIndex = ABGR1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 17;
					}
					else if (p_intSISpCd==108) //lodgepole
					{
						getAvgDbhAndBasalArea(p_intSICondId);
						p_dblSiteIndex = zPICO3(p_intSIAgeDia,
							p_intSIHtFt,
							this.ConditionClassBasalAreaPerAcre,
							this.ConditionClassAverageDia);
						p_intSIFVSSpecies=108;
					}
					else if (p_intSISpCd==93) //englemann spruce
					{
						p_dblSiteIndex = PIEN3(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=93;
					}
					else if (p_intSISpCd==19 || p_intSISpCd==22) //subalpine fir,noble fir
					{
						p_dblSiteIndex = zABPR2(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=22;
					}
					else if (p_intSISpCd==122 || p_intSISpCd==116)  //ponderosa pine
					{
						p_dblSiteIndex = PIPO3(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=122;
					}
					else if (p_intSISpCd==264) //mountain hemlock
					{
						p_dblSiteIndex = zTSME(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=264;
					}
					else
					{
						p_dblSiteIndex = 0;
						p_intSIFVSSpecies=999;
					}

				}
				//
				//Klamath Mountains variant
				//
				if (this.FVSVariant=="NC")
				{
					if (p_intSISpCd==202 || //Douglas-fir
                        p_intSISpCd==211 || //other softwoods
                        p_intSISpCd==98  ||
                        p_intSISpCd==103 ||
                        p_intSISpCd==127 ||
                        p_intSISpCd==201 ||
                        p_intSISpCd==264 ||
                        p_intSISpCd==263 || //Douglas-fir
                        p_intSISpCd==41)    
					{
						p_dblSiteIndex = qPSME1(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=202;
					}
					else if (p_intSISpCd==20 ||
						p_intSISpCd==21 ||
						p_intSISpCd==15 ||
						p_intSISpCd==81 ||
                        p_intSISpCd==17)  //red firs, white firs, incense cedar
					{
						p_dblSiteIndex = zABCO2(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=15;
					}
					else if (p_intSISpCd==361) //Madrone
					{
						p_dblSiteIndex = zARME1(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies = 361;
					}
					else if (p_intSISpCd==818) //California black oak
					{
						p_dblSiteIndex = zQUKE(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=818;
					}
					else if (p_intSISpCd==631) //tan oak
					{
						p_dblSiteIndex = zLIDE1(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=631;
					}
					else if (p_intSISpCd==117 ||    //Sugar pine 
                             p_intSISpCd==122 ||    //Ponderosa pine
                             p_intSISpCd==116 ||
                             p_intSISpCd==124 ||
                             p_intSISpCd==108 ||    //Ponderosa pine
                             p_intSISpCd==119)      //Sugar pine     	    
                             
					{
						p_dblSiteIndex = zPIPO8(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies= 122;
					}
					else
					{
						p_dblSiteIndex = 0;
						p_intSIFVSSpecies=999;
					}

				}
				//
				//South Central Oregon / Northeast California variant
				//
				if (this.FVSVariant=="SO")
				{
					if (p_intSISpCd==119) //western white pine
					{
						p_dblSiteIndex = PIMO2(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=119;
					}
					else if (p_intSISpCd==117)  //sugar pine
					{
						p_dblSiteIndex = PIPO3(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=122;  
					}
					else if (p_intSISpCd==202) //Douglas-fir
					{
						p_dblSiteIndex=qPSME12(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=202;
					}
					else if (p_intSISpCd==15)	  //white fir
					{
						p_dblSiteIndex=ABGR1(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=15;
					}
					else if (p_intSISpCd==264) //mountain hemlock
					{
						p_dblSiteIndex = zTSME(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=264;
					}
					else if (p_intSISpCd==81)	  //incense cedar
					{
						p_dblSiteIndex=ABGR1(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=15;
					}
					else if (p_intSISpCd==108) //lodgepole
					{
						getAvgDbhAndBasalArea(p_intSICondId);
						p_dblSiteIndex = zPICO3(p_intSIAgeDia,
							p_intSIHtFt,
							this.ConditionClassBasalAreaPerAcre,
							this.ConditionClassAverageDia);
						p_intSIFVSSpecies=108;
					}
					else if (p_intSISpCd==93) //englemann spruce
					{
						p_dblSiteIndex = PIEN3(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=93;
					}
					else if (p_intSISpCd==20 || p_intSISpCd==21) //red fir
					{
						p_dblSiteIndex = zABPR2(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=22;
					}
					else if (p_intSISpCd==122 || p_intSISpCd==116)  //ponderosa pine,jeffrey pine
					{
						p_dblSiteIndex = PIPO3(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=122;
					}
                    else if (p_intSISpCd == 64)  //western juniper
                    {
                        getAvgDbhAndBasalArea(p_intSICondId);
                        p_dblSiteIndex = SI_LP5(p_intSIAgeDia, p_intSIHtFt, 
                            this.ConditionClassBasalAreaPerAcre, 
                            this.ConditionClassAverageDia);
                        p_intSIFVSSpecies = 64;
                    }
					else
					{
						p_dblSiteIndex = 0;
						p_intSIFVSSpecies=999;
					}
				}
				//
				//West-side Sierra Nevada
				//
				if (this.FVSVariant=="WS")
				{
					//Adjustment factors of Dunning by species are from FVS Wessin documentation
                    //Dunning adjustment factors also listed on table 3.4.1.2 in FVS WS Variant Overview from 11/2015
                    if (p_intSISpCd == 119)		//Western white pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 119;
                    }
                    else if (p_intSISpCd == 117) //Sugar pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 117;
                    }
                    else if (p_intSISpCd == 202) //Douglas-fir
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 202;
                    }
                    else if (p_intSISpCd == 15)  //White fir
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 15;
                    }
                    else if (p_intSISpCd == 212)  //Giant sequoia
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 212;
                    }
                    else if (p_intSISpCd == 81)   //Incense cedar
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.76;
                        p_intSIFVSSpecies = 81;
                    }
                    else if (p_intSISpCd == 116)   //Jeffrey Pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 116;
                    }
                    else if (p_intSISpCd == 122)   //Ponderosa Pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 122;
                    }
                    else if (p_intSISpCd == 64)   //Western Juniper
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 64;
                    }
                    else if (p_intSISpCd == 101)   //whitebark pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 101;
                    }
                    else if (p_intSISpCd == 108)   //lodgepole pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 108;
                    }
                    else if (p_intSISpCd == 109)   //coulter pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 109;
                    }
                    else if (p_intSISpCd == 113)   //limber pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 113;
                    }
                    else if (p_intSISpCd == 120)   //bishop pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 120;
                    }
                    else if (p_intSISpCd == 127)   //California foothill pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 127;
                    }
                    else if (p_intSISpCd == 201)   //bigcone Douglas-fir
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 201;
                    }
                    else if (p_intSISpCd == 264)   //mountain hemlock
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 264;
                    }
                    else if (p_intSISpCd == 818)  //Black oak
                    {
                        p_dblSiteIndex = this.qPSME1(p_intSIAgeDia, p_intSIHtFt); //uses King's Douglas-fir
                        p_intSIFVSSpecies = 202;
                    }
                    else if (p_intSISpCd == 20)  //red fir
                    {
                        p_dblSiteIndex = this.zABMA2(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 20;
                    }
                    else if (p_intSISpCd == 631)  //tan oak
                    {
                        p_dblSiteIndex = this.qPSME1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 202;
                    }
                    else if (p_intSISpCd == 103)  //knobcone pine to great basin bristlecone pine
                    {                             //KP maps to GB for this variant according to FVS
                        p_dblSiteIndex = this.PIEN3(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 103;
                    }
                    else
                    {
                        p_dblSiteIndex = 0;
                        p_intSIFVSSpecies = 999;
                    }
				}

				if (this.FVSVariant=="CA")
				{
					//Note: this crosswalk came from a worksheet from CAheight_ref.xls sent
                    //by Chad, FVS-Fort Collins group, 4/30/2002

					if (p_intSISpCd==41  ||  p_intSISpCd==81 ||
						p_intSISpCd==242 ||  p_intSISpCd==15 ||
						p_intSISpCd==202 ||  p_intSISpCd==263 ||
						p_intSISpCd==117 ||  p_intSISpCd==92  ||
                        p_intSISpCd==212 ||  p_intSISpCd==120 ||
                        p_intSISpCd==17)
					{
						//Port Orford cedar, incense cedar, western redcedar, 
						//white-fir, Douglas-fir, western hemlock
                        //sugar pine (!), brewer spruce, giant sequoia, bishop pine

						p_dblSiteIndex = zPSME14(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies = 202;
					}
					else if (p_intSISpCd==20 || p_intSISpCd==21 ||
						p_intSISpCd==264)  //CA red fir, Shasta red fir,mountain hemlock
					{
						p_dblSiteIndex = zABMA2(p_intSIAgeDia,p_intSIHtFt);
						p_intSIFVSSpecies=20;
					}
					else if (p_intSISpCd==108 || p_intSISpCd==101 ||
						p_intSISpCd==103 || p_intSISpCd==109 || 
						p_intSISpCd==113 ||	p_intSISpCd==64) 
					{
						//Whitebark pine, knobcone pine, lodgepole pine,
						//Coulter pine, Limber pine, Western juniper
						getAvgDbhAndBasalArea(p_intSICondId);
						p_dblSiteIndex = zPICO3(p_intSIAgeDia,
							p_intSIHtFt,
							this.ConditionClassBasalAreaPerAcre,
							this.ConditionClassAverageDia);
						p_intSIFVSSpecies=108;
					}
					else if (p_intSISpCd==116 || p_intSISpCd==119 ||
						p_intSISpCd==122 || p_intSISpCd==124 ||
						p_intSISpCd==127)
					{
						//Jeffrey pine, western white pine, ponderosa pine,
						//Monterey pine, gray pine
						p_dblSiteIndex  = zPIPO9(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=122;
					}
					else if (p_intSISpCd==818 || p_intSISpCd==231 ||
						p_intSISpCd==815 || p_intSISpCd==821 || 
						p_intSISpCd==330 ||p_intSISpCd==492 ||
						p_intSISpCd==542)
					{
						//Pacific yew, Oregon white oak, California black oak, valley white oak,
						//California buckeye, Pacific dogwood, Oregon ash
						p_dblSiteIndex = zQUKE(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=818;
					}
					else if (p_intSISpCd==361 || p_intSISpCd==801 ||
						p_intSISpCd==805 || p_intSISpCd==807 ||
						p_intSISpCd==811 || p_intSISpCd==839 ||
						p_intSISpCd==431)
					{
						//Coast live oak, canyon live oak, blue oak, engelmann oak,
						//interior live oak, Pacific madrone, golden chinkapin
						p_dblSiteIndex = zARME1(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies = 361;
					}
					else if (p_intSISpCd==631) //tan oak
					{
						p_dblSiteIndex = zLIDE1(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies=631;
					}
					else if (p_intSISpCd==312 || p_intSISpCd==351 ||
						p_intSISpCd==600 || p_intSISpCd==603 ||
						p_intSISpCd==747 || p_intSISpCd==981)
					{
						//Bigleaf maple, red alder, walnuts, black cottonwood, CA laurel
						p_dblSiteIndex = zALRU3(p_intSIAgeDia, p_intSIHtFt);
						p_intSIFVSSpecies = 351;
					}
					else
					{
						p_dblSiteIndex = 0;
						p_intSIFVSSpecies=999;
					}
				}

                // Variants for ID and MT that were implemented using a site
                // index equation database
                if (this.FVSVariant == "CI" || this.FVSVariant == "EM"
                    || this.FVSVariant == "IE" || this.FVSVariant == "TT")
                {
                    // The compound key for the dictionary is the variant + species code
                    string strKey = this.FVSVariant + SI_DELIM + p_intSISpCd;
                    // Initialize values to blank in case the key is not found
                    string strEquation = "";
                    string strSlfSpCd = "";
                    string strRegion = "";
                    if (_dictSiteIdxEq.ContainsKey(strKey))
                    {
                        // If the key is found extract the values from the delimited string
                        string[] arrValues = _dictSiteIdxEq[strKey].Split(Convert.ToChar(SI_DELIM));
                        strEquation = arrValues[0];
                        strSlfSpCd = arrValues[1];
                        strRegion = arrValues[2];
                    }
                    // Reset site index and species code, in case they aren't found in database
                    p_dblSiteIndex = 0;
                    // Return the numeric FVS-FIA species code; 
                    p_intSIFVSSpecies = p_intSISpCd;
                    // Calculate the site index for the equation from the database
                    switch (strEquation)
                    {
                        case "ABGR1":
                            p_dblSiteIndex = ABGR1(p_intSIAgeDia, p_intSIHtFt);
                            break;
                        case "LAOC1_OR":
                            p_dblSiteIndex = LAOC1_OR(p_intSIAgeDia, p_intSIHtFt);
                            break;
                        case "PIEN3":
                            p_dblSiteIndex = PIEN3(p_intSIAgeDia, p_intSIHtFt);
                            break;
                        case "PSME11":
                            p_dblSiteIndex = PSME11(p_intSIAgeDia, p_intSIHtFt);
                            break;
                        case "SI_AS1":
                            p_dblSiteIndex = SI_AS1(p_intSIAgeDia, p_intSIHtFt);
                            // substitute FIA site index if 0 returned; Uses same site index equation
                            if (p_dblSiteIndex == 0)
                            {
                                p_dblSiteIndex = Convert.ToDouble(p_intSiTree);
                            }
                            break;
                        case "SI_DF2":
                            p_dblSiteIndex = SI_DF2(p_intSIAgeDia, p_intSIHtFt, this.ConditionClassHabitatTypeCd);
                            break;
                        case "SI_LP5":
                            getAvgDbhAndBasalArea(p_intSICondId);
                            p_dblSiteIndex = SI_LP5(p_intSIAgeDia, p_intSIHtFt, this.ConditionClassBasalAreaPerAcre,
                                this.ConditionClassAverageDia);
                            break;
                        case "SI_PP6":
                            p_dblSiteIndex = SI_PP6(p_intSIAgeDia, p_intSIHtFt);
                            break;
                        default:
                            p_intSIFVSSpecies = 999;
                            break;
                    }
                    // If there is a cross-reference slf species code use it, otherwise use input species code
                    int intSpCd = -1;
                    bool boolSlfSpecies = Int32.TryParse(strSlfSpCd, out intSpCd);
                    if (boolSlfSpecies) p_intSIFVSSpecies = intSpCd;
                }

					

			}
			
			/// <summary>
			///-- SITE INDEX FOR PACIFIC SILVER FIR - Hoyer
			///-- Height-age and Site Index Curves for Pacific
			///-- Silver Fir if in the Pacific Northwest
			///-- Research Paper:  PNW-RP-418 Hoyer and Herman, Sept. 1989
			///-- Curves for high elevation ABAM in West Cascades
			///-- Site index at breast high age of 100
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double ABAM2(int p_intSIDiaAge, int p_intSIHtFt)
			{
				double dblSI;
				int intHtFt;
				int intAgeDia;

				intAgeDia=p_intSIDiaAge;
				intHtFt = p_intSIHtFt;
				dblSI = intHtFt * Math.Exp((-0.0268797) * (intAgeDia - 100) / intAgeDia + 0.0046259 * Math.Pow((double)(intAgeDia - 100),2) / 100 
					- 0.0015862 * Math.Pow((double)(intAgeDia - 100),3) / 10000 
					- 0.0761453 * (intAgeDia - 100) / (Math.Pow((double)intHtFt,0.5)) 
					+ 0.0891105 * (intAgeDia - 100) / intHtFt);
				return dblSI;
			}
			/// <summary>
			///-- SITE INDEX FOR GRAND AND WHITE FIR - Cochran
			///-- Site Index and Height Growth Curves for Managed Stands of
			///-- White or Grand fir East of the Cascades in Oregon and Washinton
			///-- Research Paper:  PNW-252,  April 1979
			///-- Site index at breast high age of 50
			/// tmb note: checked some values against graph in publication
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double ABGR1(int p_intSIDiaAge, int p_intSIHtFt)
			{
				double a, b, logTreeAge;
				double dblSI;
				if (p_intSIDiaAge > 125) logTreeAge = Math.Log(125);
				else logTreeAge = Math.Log(p_intSIDiaAge);

				a = 3.886 - 1.8017 * logTreeAge + 0.2105 * (logTreeAge * logTreeAge)
					- 0.0000002885 * Math.Pow(logTreeAge, 9) + 1.187E-18
					* Math.Pow(logTreeAge, 24);

				b = -0.30935 + 1.2383 * logTreeAge + 0.001762 * Math.Pow(logTreeAge, 4)
					- 5.4E-06 * Math.Pow(logTreeAge, 9) + 2.046E-07
					* Math.Pow(logTreeAge, 11) - 4.04E-13 * Math.Pow(logTreeAge, 18);

				dblSI = (p_intSIHtFt - 4.5) * Math.Exp(a) - Math.Exp(a) * Math.Exp(b)
					+ 89.43;
				return dblSI;

			}
			/// <summary>
			///-- SITE INDEX FOR ENGELMANN SPRUCE - Clendenen/Alexander
			///-- Base-Age Conversion and Site Index Equations for Englemann
			///-- Spruce Stands in the Central and Southern Rocky Mountians
			///-- Research Note:  INT-223  1977
			///-- Clendenen equations for Alexander SI converted from 100 to 50 yr.
			///-- age = Total tree Age
			///-- si = Site index at total age of 100
 			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double PIEN3(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI;
				int intHtFt;
				double dblSI50;
				double a=0,b=0;
				int intTotalAge;
    
				if (p_intSIDiaAge + 15 > 300) intTotalAge=300;
				else intTotalAge = p_intSIDiaAge + 15;   //Arbitrary adjustment by Tara
				intHtFt = p_intSIHtFt;
				if (intTotalAge < 50)
				{
					a = 7.3214 - 0.08797 * (Math.Pow((double)(intTotalAge - 20),1.3));
					b = 2.2366 - 0.43083 * (Math.Pow((double)(intTotalAge - 20),0.31));
				}
//				else if (intTotalAge > 291) dblSI = -999;
				else
				{
					a = -25.4094 + (0.00001477) * (Math.Pow((double)(300 - intTotalAge),2.6));
					b = 0.7121 + (7.4576E-17) * (Math.Pow((double)(300 - intTotalAge), 6.5));
				}
				dblSI50= a + b * intHtFt;
				dblSI = 1.2764 * dblSI50 + 14.1943;
				return dblSI;
			}
			/// <summary>
			/// calculate red fir - Dolph 1991 PSW-RP-206 "Polymorphic site index curves for red fir 
			/// in California and southern Oregon"
			/// tmb 4/17/02  Note: Dolph uses an equation where SI can not be solved for explicitly.
			/// Hence the brute force method used below.
			/// si = site index at breast high age of 50
			/// Note: this equation as published in Dolph is missing a minus sign before 
			/// the b in the test_ht equation, and has an extra parenthesis.  
			/// Notified Martin Richie at PSW 4/16/02.
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double zABMA2(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI=0;
				double dblTest_SI;
				double b, b50, dblTest_Ht;
				double b1, b2, b3, b4, b5;
				bool bDone;
    
				b1 = 1.51744;
				b2 = 0.00000141512;
				b3 = -0.0440853;
				b4 = -3049510;
				b5 = 0.000572474;
    
				dblTest_SI = 10;
				bDone = false;
				while (!bDone)
				{
					b = p_intSIDiaAge * Math.Exp(p_intSIDiaAge * b3) * b2 * dblTest_SI + Math.Pow((double)(p_intSIDiaAge * Math.Exp(p_intSIDiaAge * b3) * b2),2) * b4 + b5;
					b50 = 50 * Math.Exp(50 * b3) * b2 * dblTest_SI + Math.Pow((double)(50 * Math.Exp(50 * b3) * b2),2) * b4 + b5;
					dblTest_Ht = ((dblTest_SI - 4.5) * (1 - Math.Exp(-b * Math.Pow((double)p_intSIDiaAge,b1)))) / (1 - Math.Exp(-b50 * Math.Pow((double)50,b1))) + 4.5;
					if (dblTest_Ht > p_intSIHtFt)
					{
						bDone = true;
						dblSI = dblTest_SI - 0.5;
					}
					else
					{
						dblTest_SI = dblTest_SI + 0.5;
					}
				}
				return dblSI;
			}
			
			/// <summary>
			///-- CALCULATE NOBLE FIR SITE INDEX - Herman, Curtis and DeMars
			///-- Height Growth and Site Index Estimates in High Elevation
			///-- Forests of the Oregon Washington Cascades   (PNW-243)
			///-- bh_age = Ring count at breast height
			///-- si = Site index at breast high age of 100
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double ABPR1(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI=0;
				double a, b, c;
				if (p_intSIDiaAge <= 100)
				{
					dblSI = 4.5
						+ 0.2145
						* (100.0 - p_intSIDiaAge)
						+ 0.0089
						* ((100.0 - p_intSIDiaAge) * (100.0 - p_intSIDiaAge))
						+ (p_intSIHtFt - 4.5)
						* (1.0 + 0.00386 * (100.0 - p_intSIDiaAge) + 1.2518
						* Math.Pow((100.0 - p_intSIDiaAge), 5) / Math.Pow(10.0, 10));
				}
				else
				{
					a = -62.755 + 672.55 * Math.Pow((1.0 / p_intSIDiaAge), 0.5);
					b = 0.9484 + 516.49 * ((1.0 / p_intSIDiaAge) * (1.0 / p_intSIDiaAge));
					c = -0.00144 + 0.1442 * (1.0 / p_intSIDiaAge);
					dblSI = a + b * (p_intSIHtFt - 4.5) + c
						* ((p_intSIHtFt - 4.5) * (p_intSIHtFt - 4.5));
				}
				return dblSI;
			}
			
			/// <summary>
			/// qPSME13 is taken from Bruce's site index routine.  Source:  Curtis, RO, Herman, FR, DeMars, DJ.
			///1974.  Height growth and site index for Douglas-fir in high elevation forests of the Oregon-
			/// Washington cascades.  Forest Science 20(4):307-316
			/// si = breast height age at 100 years
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double qPSME13(int p_intSIDiaAge, int p_intSIHtFt)
			{
				double dblSI=0;
				double a, b;
				if (p_intSIDiaAge <= 100)
				{
					a = .010006 * ((100 - p_intSIDiaAge) * (100 - p_intSIDiaAge));
					b = 1 + .00549779 * (100 - p_intSIDiaAge) + (1.46842 * Math.Pow(10, -14))
						* (Math.Pow(100 - p_intSIDiaAge, 7));
				}
				else
				{
					a = 7.66772 * (Math.Exp(-0.95 * Math.Pow(100.0 / (p_intSIDiaAge - 100.0), 2)));
					b = 1.0 - 0.730948 * Math.Pow(Math.Log10((double) p_intSIDiaAge) - 2, 0.8);

				}
				dblSI = a + b * (p_intSIHtFt - 4.5) + 4.5;
				return dblSI;
			}

			///<summary>Some site index equations require the Canopy Crown Factor (CCF)
            ///which is a function of avg dbh and basal area per plot. This function queries
            ///the TREE table to calculate these two factors for a given condition.
            ///This function needs to be called prior to running an equation that uses CCF
            ///</summary>
            ///<param name="p_intCondId">The id of the condition</param>  
            private void getAvgDbhAndBasalArea(int p_intCondId)
            {

				this.ConditionClassAverageDia=0;
                this.ConditionClassBasalAreaPerAcre = 0;

                string strSQL = "SELECT t.tpacurr, t.dia " +
                "FROM " + this.TreeTable + " t " +
                "WHERE t.biosum_cond_id = '" +
                this.BiosumPlotId + Convert.ToString(p_intCondId).Trim() +
                "' AND t.statuscd=1 AND t.tpacurr IS NOT NULL and t.dia IS NOT NULL AND t.dia >= 1";

                _oAdo.SqlQueryReader(_oAdo.m_OleDbConnection, strSQL);

                //Variables to accumulate the data from the tree table
                double dblCount = 0;
                double dblDia = 0;
                double dblBasalArea = 0;
                if (_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (_oAdo.m_OleDbDataReader.Read())
                    {
                        double tempTpa = Convert.ToDouble(_oAdo.m_OleDbDataReader["tpacurr"]);
                        double tempDia = Convert.ToDouble(_oAdo.m_OleDbDataReader["dia"]);
                        dblCount = dblCount + tempTpa;
                        dblDia = dblDia + tempDia * tempTpa;
                        dblBasalArea = dblBasalArea + (Math.Pow(tempDia, 2) * 0.00545415) * tempTpa;
                    }
                }

                if (dblCount > 0)
                {
                    this.ConditionClassAverageDia = dblDia / dblCount;
                }
                this.ConditionClassBasalAreaPerAcre = dblBasalArea;
			}

			/// <summary>
			///-- SITE INDEX FOR LOGDGEPOLE PINE - DAHMS
			///-- Gross Yield of Central Oregon Lodgepole Pine
			///-- Research Paper: PNW-8  1964
			///-- bh_age = Ring count at breast height
			///-- si = Site index at breast high age of 100
			///Tara's note: This is the same function as PICO2,
			/// but with the addition of a density correction factor from
			/// the Dahm's 1975 publication. If the p_dblBasalArea field or p_dblAvgDbh field
			/// is null or results in a CCF less than 125, no correction factor is applied.
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <param name="p_dblBasalArea"></param>
			/// <param name="p_dblAvgDbh"></param>
			/// <returns></returns>
			private double zPICO3(int p_intSIDiaAge,int p_intSIHtFt, double p_dblBasalArea, double p_dblAvgDbh)
			{
				double dblSI;
				int intHtFt;
				int intTtlAge;
				double a,b,CCF;

				intTtlAge = p_intSIDiaAge + 10;
				intHtFt = p_intSIHtFt;
				a = 68.18 - 8.8 * (Math.Pow((double)intTtlAge,0.45));
				b = 2.2614 - 1.26489 * (Math.Pow((double)(1 - Math.Exp(-0.08333 * intTtlAge)),5));
				dblSI = a + 4.5 + b * (intHtFt - 4.5);
				if (p_dblBasalArea == 0 || p_dblAvgDbh == 0)
				{
				}
				else
				{
					CCF = Math.Exp(1.25021 + 0.97834 * Math.Log10(p_dblBasalArea) - 
						0.524548 * Math.Log10(p_dblAvgDbh)); //p.218 of Dahms(1975)
					if ((CCF > 125) && (p_intSIDiaAge > 41))
						dblSI = dblSI / (1 - 0.00082 * (CCF - 125)); //corrected for density
        
				}
				return dblSI;
			}
			/// <summary>
			///-- SITE INDEX FOR WESTERN WHITE PINE - Curtis
			///-- PNW-RP-423 (Cascades - Mt Hood and Gifford Pinchot NF's)
			///-- Height Growth and Site Index for Western White Pine in the
			///-- Cascade Range of Washington and Oregon
			///-- IAGE = Ring count at breast height
			///-- si = Site index at breast high age of 100
			/// tara note: checked against figures in publication
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double PIMO3(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI;
				int intHtFt;
				int intDiaAge;
				double c1;
				double a,b;
   
				intDiaAge = p_intSIDiaAge;
				intHtFt = p_intSIHtFt;
				c1 = intDiaAge - 100;
				a = Math.Exp(0.37072 * (Math.Log((double)intDiaAge) - Math.Log((double)100)) - 0.03745 * c1 + (0.00021645) * Math.Pow((double)c1,2));
				b = (Math.Pow((double)intHtFt,(1 + 0.005936 * c1 - (0.00003879) * Math.Pow((double)c1,2))));
				dblSI = a * b;
				return dblSI;
			}
			/// <summary>
			///-- SITE INDEX FOR PONDEROSA PINE - Barrett
			///-- Height Growth and Site Index Curves for Managed, Even-aged
			///-- Stands of Ponderosa Pine in the Pacific Northwest
			///-- Reseach Paper:  PNW-232
			///-- bh_age = ring count at 4.5 feet
			///-- si = site index at breast high age of 100
			/// Tara note: checked some points against figure in publication - seems o.k.
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double PIPO3(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI=0;
				double a, b, c;
				if (p_intSIDiaAge <= 130)
				{
					a = 100.43;
					b = 1.198632 - (0.00283073 * p_intSIDiaAge) + (8.44441 / p_intSIDiaAge);
					c = 128.8952205 * (Math.Pow(1.0 - (Math.Exp(-0.016959 * p_intSIDiaAge)),
						1.23114));
					dblSI = a - (b * c) + (b * (p_intSIHtFt-4.5)) + 4.5;
				}
				else
				{
					dblSI = ((5.328 * (Math.Pow(p_intSIDiaAge, -0.1)) - 2.378) * (p_intSIHtFt - 4.5)) + 4.5;
				}
				return dblSI;
			}
			/// <summary>
			///-- SITE INDEX FOR WESTERN HEMLOCK - Wiley
			///-- WEYERHAUSER FORESTRY PAPER NO.17 - PAGE 10, 1978
			///-- Site Index Tables for Western Hemlock in the Pacific Northwest
			///-- age = Ring count at breast height
			///-- si = Site index at breast high age of 50
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double TSHE1(int p_intSIDiaAge, int p_intSIHtFt)
			{
				double denom;
				double dblSI=0;

				if (p_intSIDiaAge <= 120)
				{
					denom = ((p_intSIDiaAge * p_intSIDiaAge) - (p_intSIHtFt - 4.5)
						* (-1.7307 - 0.0616 * p_intSIDiaAge + 0.00192 * (p_intSIDiaAge * p_intSIDiaAge)));
					if (denom != 0)
					{
						dblSI = 2500 * (((p_intSIHtFt - 4.5) * (0.1394 + 0.0137 * p_intSIDiaAge + 0.00007 * (p_intSIDiaAge * p_intSIDiaAge))) / denom) + 4.5;
					}

				}
				else
				{
					dblSI = 4.5 + 22.6 * Math.Exp((0.014482 - 0.001162 * Math.Log(p_intSIDiaAge))
						* (p_intSIHtFt - 4.5));
				}
				return dblSI;
			}
			/// <summary>
			///calculate Means mountain hemlock site index.  From "Means, Joseph E.,
			/// Mary H. Campbell, and Gregory P. Johnson.  1988.  FIR report Vol. 10, No. 1,
			/// p8-9."  Also unpublished draft manuscript - has more detail.
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double zTSME(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI;
				double a,b,c,d,e,f,g;
				double dblMetricHeight;
				a = 17.22;
				b = 0.583224;
				c = 99.1275;
				d = -1.189892;
				e = 47.9256;
				f = -0.00574787;
				g = 1.241599;
        
				dblMetricHeight = p_intSIHtFt * 0.3048; //Convert from feet to metric
				dblSI = 1.37 + a + (b + c * Math.Pow(p_intSIDiaAge,d)) * (dblMetricHeight - 1.37 - e * Math.Pow((double)(1 - Math.Exp(f * p_intSIDiaAge)),g));

				dblSI = dblSI / 0.3048; //Convert from metric to feet
				return dblSI;
			}
			/// <summary>
			///-- CALCULATE RED ALDER - Worthington
			///-- (PNW-36) - Normal Yeild Tables for Red Alder - 1960
			///-- bh_age = Ring count at breast height
			///-- si = Site index at breast high age of 50
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double ALRU2(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI = (0.60924 + 19.538 / p_intSIDiaAge) * p_intSIHtFt;
				return dblSI;
			}
			
			/// <summary>
			/// - SITE INDEX FOR WESTERN LARCH - COCHRAN
			///-- Site Index, Height Growth, Normal Yields, and Stocking
			///-- Levels for Larch in Oregon and Washinton
			///-- Research Note: PNW-424, May 1985
			///-- bh_age = Ring count at breast height
			///-- tht = Total tree height
			///-- si = Site index at breast high age of 50
			/// </summary>
			/// <param name="p_intSIDiaAge">ring count at breast height</param>
			/// <param name="p_intSIHtFt">total tree height in feet</param>
			/// <returns></returns>
			private double LAOC1_OR(int p_intSIDiaAge,int p_intSIHtFt)
			{
				int intSIDiaAge=0;
				if (p_intSIDiaAge > 100) intSIDiaAge = 100;
				else intSIDiaAge=p_intSIDiaAge;

				double dblSI = 78.07
					+ (p_intSIHtFt - 4.5)
					* (3.51412 - 0.125483 * intSIDiaAge + 0.0023559 * Math.Pow(intSIDiaAge, 2)
					- 0.00002028 * Math.Pow(intSIDiaAge, 3) + 0.000000064782 * Math.Pow(
					intSIDiaAge, 4))
					- (3.51412 - 0.125483 * intSIDiaAge + 0.0023559 * Math.Pow(intSIDiaAge, 2)
					- 0.00002028 * Math.Pow(intSIDiaAge, 3) + 0.000000064782 * Math.Pow(
					intSIDiaAge, 4))
					* (1.46897 * intSIDiaAge + 0.0092466 * Math.Pow(intSIDiaAge, 2) - 0.00023957
					* Math.Pow(intSIDiaAge, 3) + 0.0000011122 * Math.Pow(intSIDiaAge, 4));
				return dblSI;
			}
			/// <summary>
			/// western larch in western washington and eastern washington;
			/// and western white pine in eastern washington (Brickell 1970)
			/// This equation replaced Cochran (1985) equation for the same areas and species
			/// beginning in 1990
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double LAOC1_WA(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI = 0.37956 * p_intSIHtFt * Math.Exp(48.4372 / (p_intSIDiaAge + 8));
				return dblSI;
			}
			/// <summary>
			///-- SITE INDEX FOR WESTERN WHITE PINE - Brickell and Haig
			///-- Equations and computer subroutines for Estimating Site
			///-- Quality of Eight Rocky Mountian Species
			///-- Reaserch Paper: INT-75  1970
			///-- age = Age of dominant trees in even-aged stand
			///-- tht = Average height of dominant and codominant trees
			///-- si = Site index at base age of 50
			/// Tara notes: Apparently developed from Haig, Irvine T. 1932.  Second-growth yield, stand and volume
			/// tables for the western white pine type.  USDA Tech. Bull. 323.
			/// Probably total age?  Note non-standard selection of "height" (average) - this function operates on
			/// individual trees.
			/// Changed coefficients to match Brickell publication
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double PIMO2(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI;
				int intHtFt;
				int intDiaAge;
				intDiaAge = p_intSIDiaAge + 11; //Average reported in Haig 1932.
				intHtFt = p_intSIHtFt;
				dblSI = 0.37504453 * intHtFt * (Math.Pow((double)(1 - 0.92503 * Math.Exp(-0.0207959 * p_intSIDiaAge)),-2.4881068));
				return dblSI;
			}
			/// <summary>
			///calculate Cochran Douglas-fir index.  From "Cochran, PH. 1979.
			///Site index and height growth curves for
			///managed, even-aged stands of Douglas-fir east of the Cascades in Oregon and Washington.
			///RP-PNW-251."
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double qPSME12(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI;
				double a, b, logTreeAge;
				
				if (p_intSIDiaAge > 100) logTreeAge = Math.Log(100);
				else logTreeAge = Math.Log(p_intSIDiaAge);
				a = Math.Exp(-0.37496 + 1.36164 * logTreeAge - 0.00243434
					* Math.Pow((double) logTreeAge, 4));
				b = 0.52032 - 0.0013194 * p_intSIDiaAge + 27.2823 / p_intSIDiaAge;
				dblSI=84.47 - a * b + (p_intSIHtFt - 4.5) * b;
				return dblSI;
			
			}
			
			/// <summary>
			///-- Site Index Noble fir - DeMars, D.J., Herman, F.R., and J.F. Bell.  1970.  
			///Preliminary site index curves for noble fir from stem analysis data.
			///-- PNW Research Note PNW-119 May 1970.  si = 100 years at breast height age, total height
			///-- Based on tallest dominant
			///-- Tara's note: the original publication contained graphs and a table,
			/// developed from 4 polymorphic curves
			///-- This program is my adaption of the method used.
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double zABPR2(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI=0;
				double ht_A, ht_B, ht_C, ht_D;
				double propDiff;
				int intAgeDia;
				intAgeDia = p_intSIDiaAge;
    
				ht_A = ((intAgeDia * intAgeDia) / (13.64781 + (0.19937 * intAgeDia) + (0.00416 * (intAgeDia * intAgeDia)))) + 4.5;
				ht_B = ((intAgeDia * intAgeDia) / (10.11508 + (0.40115 * intAgeDia) + (0.00426 * (intAgeDia * intAgeDia)))) + 4.5;
				ht_C = ((intAgeDia * intAgeDia) / (19.10187 + (0.50996 * intAgeDia) + (0.0049 *  (intAgeDia * intAgeDia)))) + 4.5;
				ht_D = ((intAgeDia * intAgeDia) / (12.68825 + (1.05408 * intAgeDia) + (0.00501 * (intAgeDia * intAgeDia)))) + 4.5;

				if (p_intSIHtFt >= ht_A)
					dblSI = (p_intSIHtFt / ht_A) * 137.5056;  //137 is the SI for polymorphic curve A

				if ((p_intSIHtFt >= ht_B) && (p_intSIHtFt < ht_A))
				{
					propDiff = (p_intSIHtFt - ht_B) / (ht_A - ht_B);
					dblSI = (propDiff * (137.5056 - 112.2237)) + 112.2237; //112 is the SI for polymorphic curve B
				}
				if ((p_intSIHtFt >= ht_C) && (p_intSIHtFt < ht_B))
				{
					propDiff = (p_intSIHtFt - ht_C) / (ht_B - ht_C);
					dblSI = (propDiff * (112.2237 - 88.46456)) + 88.46456;  //88 is the SI for polymorphic curve B
				}
				if ((p_intSIHtFt >= ht_D) && (p_intSIHtFt < ht_C))
				{
					propDiff = (p_intSIHtFt - ht_D) / (ht_C - ht_D);
					dblSI = (propDiff * (88.46456 - 63.95436)) + 63.95436;  //63 is the SI for polymorphic curve B
				}
				if (p_intSIHtFt < ht_D) 
					dblSI = (p_intSIHtFt / ht_D) * 63.95436;

				return dblSI;
			}
			/// <summary>
			///-- SITKA SPRUCE SITE INDEX - Farr
			///-- (PNW-326) Site Index and Height Growth Curves for Unmanaged Even-
			///-- Aged Stannds of Western Hemlock and Sitka Spruce In Southeast Alaska
			///-- Research Paper: PNW-326,  OCT. 1984
			///-- intAgeDia = Ring count at breast height
			///-- si = Site index at breast high age of 50
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double PISI1(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI;
				int intHtFt, intDiaAge;
				double c1, c2;

				intDiaAge = p_intSIDiaAge;
				intHtFt = p_intSIHtFt;
				c1 = 6.396816 - 4.098921 * Math.Log(intDiaAge) + 0.76287 * (Math.Pow((double)(Math.Log(intDiaAge)),2))
					- (0.00244688) * (Math.Pow((double)(Math.Log(intDiaAge)),5)) 
					+ (0.000000244581) * (Math.Pow((double)(Math.Log(intDiaAge)),10)) 
					- (2.02215E-22) * (Math.Pow((double)(Math.Log(intDiaAge)),30));
				c2 = -0.20505 + 1.4496 * Math.Log(intDiaAge) - 0.01781 * (Math.Pow((double)(Math.Log(intDiaAge)),3)) 
					+ (0.0000651975) * (Math.Pow((double)(Math.Log(intDiaAge)),5)) 
					- (1.09559E-23) * (Math.Pow((double)(Math.Log(intDiaAge)),30));
				dblSI = 90.93 - Math.Exp(c1) * Math.Exp(c2) + Math.Exp(c1) * (intHtFt - 4.5);
				return dblSI;
			}
			/// <summary>
			///-- CALCULATE DOUGLAS FIR SITE INDEX - King
			///-- Site Index for Douglas Fir in the Pacific Northwest
			///-- WEYERHAUSER FORESTRY PAPER NO. 8,  July 1966
			///-- site for breast height age 50
			///-- Taras note:  This equation is listed incorrectly in Ericas publication PNW-RN-533.  However,
			///-- the version that appears here looks correct.
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double qPSME1(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI=0;
				double z = 0, denom;
				denom = (0.109757 + (0.00792236 * p_intSIDiaAge) + 0.000197693 * (p_intSIDiaAge * p_intSIDiaAge));
				if (denom != 0)
				{
					z = ((p_intSIDiaAge * p_intSIDiaAge) / (p_intSIHtFt - 4.5) + 0.954038
						- (0.0558178 * p_intSIDiaAge) + 0.000733819 * (p_intSIDiaAge * p_intSIDiaAge))
						/ denom;

				}
				if (z != 0)
				{
					dblSI = (2500 / z) + 4.5;
				}
				return dblSI;
			}
			/// <summary>
			///-- SITE INDEX RED ALDER - Harrington
			///-- Height Growth and Site Index Curves for Red Alder
			///-- Research Paper:  PNW-326 1986
			///-- bh_age = Ring count at breast height
			///-- si = Site index at breast high age of 20
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double ALRU1(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI;
				int intHtFt,intDiaAge;
				intDiaAge = p_intSIDiaAge;
				intHtFt = p_intSIHtFt;
				dblSI = 54.185 - 4.61694 * intDiaAge + 0.11065 * intDiaAge * intDiaAge
					- 0.0007633 * intDiaAge * intDiaAge * intDiaAge
					+ (1.25934 - 0.012989 * intDiaAge + 3.522 * (Math.Pow((double)(1 / intDiaAge),3))) * intHtFt;

				return dblSI;
 			}
			/// <summary>
			/// This function was taken from Dolph, Leroy K.  1987.  Site index curves for
			//  young-growth California white fir on the western slope of the Sierra Nevada.
			//  Res. Paper PSW-185.  Berkeley, CA: PSW Forest and Range Experiment Station.
			// uses breast height age and total tree height
			//  Checked the function against tables in the publication - o.k.
			//  si = total height at 50 years bhage
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double zABCO2(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI;
				dblSI  = 69.91 - (38.020235 * (Math.Pow((double)p_intSIDiaAge,-1.052133)) * 
					(Math.Exp(p_intSIDiaAge * 0.009557)) * 
					101.842894 * (1 - Math.Exp(-0.001442 * (Math.Pow((double)p_intSIDiaAge,1.679259))))) +
					(p_intSIHtFt - 4.5) * (38.020235 * (Math.Pow((double)p_intSIDiaAge,-1.052133)) * 
					(Math.Exp(0.0095557 * p_intSIDiaAge)));
				return dblSI;
			}
			/// <summary>
			///This equation for Madrone is from "Porter, D.R. and Wiant, H.V. 1965. 
			///  Site index equations for tanoak, Pacific
			///madrone, and red alder in the redwod region of humboldt county, California.  Journal
			///of Forestry, April, 1965,p.286-287."  SI is total age at 50 years. Correction from bhage
			/// to total age is based on average values from the publication.
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double zARME1(int p_intSIDiaAge,int p_intSIHtFt)
			{

				double dblSI;
				double Total_Age;

				Total_Age = p_intSIDiaAge + 2.8 ; //Average value from study
				dblSI = (0.375 + (31.233 / Total_Age)) * p_intSIHtFt;
				return dblSI;
			}
			/// <summary>
			///Taken from "Powers, Robert F.  1972.  Site index curves for unmanaged stands of
			/// California black oak.  Research Note PSW-262.  5 p.
			/// Uses total height, breast height age, tree selection "Dominants", development method
			/// no sectioning: stratification into 5 ponderosa pine site classes, separate fitting for each
			/// resulting in 4 curves listed below.  Interpolation is my addition.  One curve discarded
			/// (Ponderosa pine SI group 110-115).  Equation given in publication doesn't match graph, appears
			/// to be incorrect (?).  Don't use for age less than 40 (curves cross).  
			/// Assumes anamorphic curves for
			/// everything that is not between si 42 and si 56.  Resulting function is slightly different than
			/// what is shown in publication figure (5.5 ft' higher at 100 yrs for SI 70, 2.5' lower
			/// at 100 years for SI 30)
			/// si = total height at breast height age of 50.
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double zQUKE(int p_intSIDiaAge,int p_intSIHtFt)
			{

				double dblSI=0;
				double htPPSI_65, htPPSI_75, htPPSI_85, htPPSI_95;
				double propDiff;

				htPPSI_65 = (Math.Sqrt(p_intSIDiaAge) - 0.442) / 0.158;  // si at 50 = 41.956
				htPPSI_75 = (Math.Sqrt(p_intSIDiaAge) - 1.888) / 0.117;  // si at 50 = 44.300
				htPPSI_85 = (Math.Sqrt(p_intSIDiaAge) - 2.267) / 0.093;  // si at 50 = 51.657
				htPPSI_95 = (Math.Sqrt(p_intSIDiaAge) - 2.08) / 0.088;   // si at 50 = 56.717

				if (p_intSIHtFt >= htPPSI_95)
					dblSI = (p_intSIHtFt / htPPSI_95) * 56.717;

				if ((p_intSIHtFt < htPPSI_95) && (p_intSIHtFt >= htPPSI_85))
				{
					propDiff = (p_intSIHtFt - htPPSI_85) / (htPPSI_95 - htPPSI_85);
					dblSI= propDiff * (56.717 - 51.657) + 51.657;
				}
				if ((p_intSIHtFt < htPPSI_85) && (p_intSIHtFt >= htPPSI_75))
				{
					propDiff = (p_intSIHtFt - htPPSI_75) / (htPPSI_85 - htPPSI_75);
					dblSI = propDiff * (51.657 - 44.3) + 44.3;
				}
				if ((p_intSIHtFt < htPPSI_75) && (p_intSIHtFt >= htPPSI_65))
				{
					propDiff = (p_intSIHtFt - htPPSI_65) / (htPPSI_75 - htPPSI_65);
					dblSI = propDiff * (44.3 - 41.956) + 41.956;
				}
				if (p_intSIHtFt < htPPSI_65) 
					dblSI = (p_intSIHtFt / htPPSI_65) * 41.956;

				return dblSI;

			}
			/// <summary>
			///This equation for Tanoak is from "Porter, D.R. and Wiant, H.V. 1965. 
			///Site index equations for tanoak, Pacific
			///madrone, and red alder in the redwod region of humboldt county, California.  Journal
			///of Forestry, April, 1965,p.286-287."  SI is total age at 50 years. Correction from bhage
			/// to total age is based on average values from the publication.
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double zLIDE1(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI;
				double Total_Age;

				Total_Age = p_intSIDiaAge + 3.2; //Average value from study
				dblSI = (0.204 + (39.787 / Total_Age)) * p_intSIHtFt;
				return dblSI;
			}
			/// <summary>
			///From Powers, Robert F. and William W. Oliver.  1978.  Site classification of 
			///ponderosa pine stands
			///stocking control in California. USDA Forest Service research paper PSW-128.
			///Powers & Oliver used total age - added 9 years to make the adjustment from bhage.
			///si = total height at 50 years total age
			///Tested function against graphed values.  Do not use for total age less than 10 years 
			///or greater than 80 yrs.
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double zPIPO8(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI=0;
				double test_Ht, test_SI;
				bool bDone;
				int Total_Age;

				Total_Age = p_intSIDiaAge + 9; //My arbitrary choice to adjust age
				test_SI = 10;
				bDone = false;
				while (!bDone)
				{
      
					test_Ht = ((1.88 * test_SI) - 7.178) * 
						(Math.Pow((double)(1 - Math.Exp(-0.025 * Total_Age)),(0.001 * test_SI + 1.64)));
					if (test_Ht > p_intSIHtFt)
					{
						dblSI = test_SI - 0.5;
						bDone = true;
					}
					else
						test_SI = test_SI + 0.5;
      
				}
				return dblSI;
			}
			/// <summary>
			///Tara note:  The following function was taken from the Region 5 "FIA User's guide" for the
			///1990s inventory on FS land.  It is described as an adaption of "Dunning, Duncan. 1942 (amended
			///1958).  Forest Research note 28, California Forest and Range Experiment Station."
			///Per Mike Landram, this adaption was done by Jack Levitan in the 1970s and has been historically
			///used by R5 since then.  For the Wessin variant development, they added 5 years to adjust for
			/// Dunning's use of stump age.  I chose to include this 5 year adaption in the function below.  The
			/// results are different than the original Dunning.  Other slight alteration from the
			/// R5 code:  SI values are calculated using interpolation, rather than assigning a site class 
			/// to each tree.
			/// This function assumes zage is breast height age
			/// si is total height at 50 years bhage.
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double zDunning(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double Ht_Site_0, Ht_Site_1, Ht_Site_2, Ht_Site_3, Ht_Site_4, Ht_Site_5;
				double dblSI=0;
				double totAge;
				double propDiff;

				totAge = p_intSIDiaAge + 5;

				Ht_Site_0 = -88.9 + 49.7067 * Math.Log(totAge);  //si @ 50 is 110.3
				Ht_Site_1 = -82.2 + 44.1147 * Math.Log(totAge);  //si @ 50 is 94.6
				Ht_Site_2 = -78.3 + 39.1441 * Math.Log(totAge);  //si @ 50 is 78.6
				Ht_Site_3 = -82.1 + 35.416 * Math.Log(totAge);   //si @ 50 is 59.8
				Ht_Site_4 = -56 + 26.7173 * Math.Log(totAge);    //si @ 50 is 51.1
				Ht_Site_5 = -33.8 + 18.64 * Math.Log(totAge);    //si @ 50 is 40.9

				if (p_intSIHtFt >= Ht_Site_0)
					dblSI = (p_intSIHtFt / Ht_Site_0) * 110.3;

				if ((p_intSIHtFt < Ht_Site_0) && (p_intSIHtFt >= Ht_Site_1)) 
				{
					propDiff = (p_intSIHtFt - Ht_Site_1) / (Ht_Site_0 - Ht_Site_1);
					dblSI = propDiff * (110.3 - 94.6) + 94.6;
				}
				if ((p_intSIHtFt < Ht_Site_1) && (p_intSIHtFt >= Ht_Site_2))
				{
					propDiff = (p_intSIHtFt - Ht_Site_2) / (Ht_Site_1 - Ht_Site_2);
					dblSI = propDiff * (94.6 - 78.6) + 78.6;
				}
				if ((p_intSIHtFt < Ht_Site_2) && (p_intSIHtFt >= Ht_Site_3))
				{
					propDiff = (p_intSIHtFt - Ht_Site_3) / (Ht_Site_2 - Ht_Site_3);
					dblSI = propDiff * (78.6 - 59.8) + 59.8;
				}
				if ((p_intSIHtFt < Ht_Site_3) && (p_intSIHtFt >= Ht_Site_4))
				{
					propDiff = (p_intSIHtFt - Ht_Site_4) / (Ht_Site_3 - Ht_Site_4);
					dblSI = propDiff * (59.8 - 51.1) + 51.1;
				}
				if ((p_intSIHtFt < Ht_Site_4) && (p_intSIHtFt >= Ht_Site_5))
				{
					propDiff = (p_intSIHtFt - Ht_Site_5) / (Ht_Site_4 - Ht_Site_5);
					dblSI = propDiff * (51.1 - 40.9) + 40.9;
				}
				if (p_intSIHtFt < Ht_Site_5)
				{
					dblSI = (p_intSIHtFt / Ht_Site_5) * 40.9;
				}
				return dblSI;
			}
			/// <summary>
			///From Hann, David W. and John A. Scrivani.  1987. Dominant-height-growth and site-index
			/// equations for Douglas-fir and ponderosa pine in Southwest Oregon.  Forest research
			/// lab, OSU, Research bulletin 59, February 1987.
			/// Program is just a duplicate of their Appendix 2.
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double zPSME14(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI;
				double temp;
				double a0, a1, a2, b1, b2, b3, b4;
				double theTest, theB;
				a0 = -6.21693;
				a1 = 0.281176;
				a2 = 1.14354;
				b1 = -0.0521778;
				b2 = 0.000715141;
				b3 = 0.00797252;
				b4 = -0.000133377;

				temp = b1 * (p_intSIDiaAge - 50) + b2 * 
					Math.Pow((double)(p_intSIDiaAge - 50),2) + b3 * Math.Log(p_intSIHtFt - 4.5) *
					(p_intSIDiaAge - 50) + b4 * Math.Log(p_intSIHtFt - 4.5) * Math.Pow((double)(p_intSIDiaAge - 50),2);
				dblSI = 4.5 + (p_intSIHtFt - 4.5) * Math.Exp(temp);
				theTest = 0;
				while (theTest < 0.999)
				{
					temp = dblSI;
					theB = 1 - Math.Exp(-Math.Exp(a0 + a1 * Math.Log(dblSI - 4.5) + a2 * 3.912023));
					theB = theB / (1 - Math.Exp(-Math.Exp(a0 + a1 * Math.Log(dblSI - 4.5) + a2 * Math.Log(p_intSIDiaAge))));
					dblSI = 4.5 + (p_intSIHtFt - 4.5) * theB;
					theTest = 1 - Math.Abs((dblSI - temp) / temp);
				}
				return dblSI; 
			}
			/// <summary>
			///From Hann, David W. and John A. Scrivani.  1987. Dominant-height-growth and site-index
			/// equations for Douglas-fir and ponderosa pine in Southwest Oregon.  Forest research
			/// lab, OSU, Research bulletin 59, February 1987.
			/// Program is just a duplicate of their Appendix 2.
			/// Note:  the publication is missing a "-EXP" in the Appendix 2 program for PP.
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double zPIPO9(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI;
				double temp;
				double a0, a1, a2, b1, b2, b3, b4;
				double theTest, theB;
				a0 = -6.54707;
				a1 = 0.288169;
				a2 = 1.21297;
				b1 = -0.069934;
				b2 = 0.000359644;
				b3 = 0.0120483;
				b4 = -0.0000718058;

				temp = b1 * (p_intSIDiaAge - 50) + b2 * 
					Math.Pow((double)(p_intSIDiaAge - 50),2) + b3 * 
					Math.Log(p_intSIHtFt - 4.5) *
					(p_intSIDiaAge - 50) + b4 * 
					Math.Log(p_intSIHtFt - 4.5) * 
					Math.Pow((double)(p_intSIDiaAge - 50),2);
     
				dblSI = 4.5 + (p_intSIHtFt - 4.5) * Math.Exp(temp);
				theTest = 0;
				while (theTest < 0.999)
				{
					temp = dblSI;
					theB = 1 - Math.Exp(-Math.Exp(a0 + a1 * Math.Log(dblSI - 4.5) + a2 * 3.912023));
					theB = theB / (1 - Math.Exp(-Math.Exp(a0 + a1 * Math.Log(dblSI - 4.5) + a2 * Math.Log(p_intSIDiaAge))));
					dblSI = 4.5 + (p_intSIHtFt - 4.5) * theB;
					theTest = 1 - Math.Abs((dblSI - temp) / temp);
				}

				return dblSI;
			}
			/// <summary>
			/// This equation for red alder is from "Porter, D.R. and Wiant, H.V. 1965. 
			/// Site index equations for tanoak, Pacific
			///madrone, and red alder in the redwod region of humboldt county, California.  Journal
			///of Forestry, April, 1965,p.286-287."  SI is total age at 50 years. Correction from bhage
			/// to total age is based on average values from the publication.
			/// </summary>
			/// <param name="p_intSIDiaAge"></param>
			/// <param name="p_intSIHtFt"></param>
			/// <returns></returns>
			private double zALRU3(int p_intSIDiaAge,int p_intSIHtFt)
			{
				double dblSI;
				double Total_Age;
				Total_Age = p_intSIDiaAge + 1.2;   //Average value from study
				dblSI = (0.649 + (17.556 / Total_Age)) * p_intSIHtFt;
				return dblSI;
			}

            /// <summary>
            /// This equation for Douglas Fir is from Brickell and Haig
            ///-- Equations and computer subroutines for Estimating Site
            ///-- Quality of Eight Rocky Mountain Species
            ///-- Research Paper: INT-75  1970. p_intSIDiaAge is total age
            ///---Derived from Kurt's PL/SQL
            ///---Curves validated by jsfried
            ///<param name="p_intSIDiaAge"></param>
            ///<param name="p_intSIHtFt"></param>
            /// </summary>
            private double PSME11(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = 0;
                double b;
                int Total_Age = p_intSIDiaAge;
                if (p_intSIDiaAge < 20)
                {
                    return dblSI;
                }
                else if (p_intSIDiaAge > 200)
                {
                    Total_Age = 200;
                }
                    else
                {
                    Total_Age = p_intSIDiaAge;
                }

                b = 40.984664 * (Math.Pow(Total_Age, -0.5) - Math.Pow(50, -0.5))
                    + 4521.1527 * (Math.Pow(Total_Age, -2.5) - Math.Pow(50, -2.5))
                    + 123059.38 * (Math.Pow(Total_Age, -3.5) - Math.Pow(50, -3.5))
                    - 0.53328682E-08 * (Math.Pow(Total_Age, 4) - Math.Pow(50, 4))
                    + 0.37808033E-10 * (Math.Pow(Total_Age, 5) - Math.Pow(50, 5))
                    + 216.64152 * p_intSIHtFt * (Math.Pow(Total_Age, -1.5) - Math.Pow(50, -1.5))
                    - 158121.49 * p_intSIHtFt * (Math.Pow(Total_Age, -4) - Math.Pow(50, -4))
                    + 1894030.8 * p_intSIHtFt * (Math.Pow(Total_Age, -5) - Math.Pow(50, -5))
                    - 0.10230592E-09 * p_intSIHtFt * (Math.Pow(Total_Age, 4) - Math.Pow(50, 4))
                    - 6.0686119 * p_intSIHtFt * p_intSIHtFt * (Math.Pow(Total_Age, -2) - Math.Pow(50, -2))
                    - 25351.090 * p_intSIHtFt * p_intSIHtFt * (Math.Pow(Total_Age, -5) - Math.Pow(50, -5))
                    + 0.33512858E-04 * p_intSIHtFt * p_intSIHtFt * (Total_Age - 50)
                    + 0.17024711E-02 * Math.Pow(p_intSIHtFt, 3) * (Math.Pow(Total_Age, -1) - Math.Pow(50, -1))
                    + 398.36720 * Math.Pow(p_intSIHtFt, 3) * (Math.Pow(Total_Age, -5) - Math.Pow(50, -5))
                    - 0.88665409E-08 * Math.Pow(p_intSIHtFt, 3) * (Math.Pow(Total_Age, 1.5) - Math.Pow(50, 1.5))
                    + 0.40019102E-14 * Math.Pow(p_intSIHtFt, 3) * (Math.Pow(Total_Age, 4) - Math.Pow(50, 4))
                    - 0.46929245E-08 * Math.Pow(p_intSIHtFt, 5) * (Math.Pow(Total_Age, -0.5) - Math.Pow(50, -0.5))
                    - 0.16640659E-20 * Math.Pow(p_intSIHtFt, 5) * (Math.Pow(Total_Age, 4.5) - Math.Pow(50, 4.5));
                dblSI = p_intSIHtFt + b;
                if (dblSI < 10)
                {
                    dblSI = 10;
                }
                else if (dblSI > 110)
                {
                    dblSI = 110;
                }
                return dblSI;
            }

            /// <summary>
            /// SITE INDEX FOR LODGEPOLE PINE - Alexander, R. R., Trackle, D. and Dahms, W. G. 1967
            /// Site indexes for lodgepole pine with corrections for stand density: methodology
            /// USDA Forest Service, Res. Pap. RM-29
            /// Site index at total age of 100
            /// Derived from VBA source code by Don Vandendriese for FIA2FVS from RMRS
            /// Curves validated by jsfried
            /// </summary>
            /// <param name="p_intSIDiaAge">Age of site tree (Ring count at breast height)</param>
            /// <param name="p_intSIHtFt">Diameter of site tree</param>
            /// <param name="p_dblAvgDbh">Average DBH per condition</param>
            /// <param name="p_dblBasalArea">Basal area per acre for the condition</param>
            /// <returns>Site Index at breast high age of 100</returns>
            private double SI_LP5(int p_intSIDiaAge, int p_intSIHtFt, double p_dblBasalArea, double p_dblAvgDbh)
            {
                double dblSI = 0;
                int Total_Age = p_intSIDiaAge;
                if (Total_Age > 0)
                {
                    Total_Age = Total_Age + 12;
                }
                else
                {
                    return dblSI;
                }

                double dblCCF = 0;
                if (p_dblAvgDbh > 0 && p_dblBasalArea > 0)
                    dblCCF = 50.58 + 5.25 * (p_dblBasalArea / p_dblAvgDbh);
                if (dblCCF > 125)
                    dblCCF = dblCCF - 125;

                dblSI = (p_intSIHtFt - 9.89331 + 0.19177 * Total_Age - 0.00124 * Math.Pow(Total_Age, 2)) / (-0.00082 * dblCCF + 0.01387 * Total_Age - 0.0000455 * Math.Pow(Total_Age, 2));
                return dblSI;
            }

            /// <summary>
            /// SITE INDEX FOR ASPEN - Edminister
            /// Site Index Curves for Aspen in the Central Rocky Mountains
            /// Research Note: RM-453
            /// Derived from VBA source code by Don Vandendriese for FIA2FVS from RMRS
            /// Base age is 80 years
            /// Curves validated by jsfried
            /// </summary>
            /// <param name="p_intSIDiaAge">Age of site tree (Ring count at breast height)</param>
            /// <param name="p_intSIHtFt">Diameter of site tree</param>
            private double SI_AS1(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = 0;
                dblSI = 4.5 + 0.48274 * (p_intSIHtFt - 4.5) * Math.Pow((1 - Math.Exp(-0.007719 * p_intSIDiaAge)), -0.93972);
                return dblSI;
            }


            /// <summary>
            /// SITE INDEX FOR PONDEROSA PINE - MEYER
		    /// Forest Service Inventory procedures for Eastern Washington,
            /// (Approximates Meyer in USDA Tech. bull. 630)
            /// Derived from VBA source code by Don Vandendriese for FIA2FVS from RMRS
            /// Base age is 100 years
            /// Site index curves validated by jsfried
            /// </summary>
            /// <param name="p_intSIDiaAge">Age of site tree (Ring count at breast height)</param>
            /// <param name="p_intSIHtFt">Diameter of site tree</param>
            private double SI_PP6(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = 0;
                int Total_Age = p_intSIDiaAge;
                // Adjustment because this equation uses total tree age
                if (Total_Age > 0) Total_Age = Total_Age + 9;
                dblSI = (5.328 * (Math.Pow(Total_Age, -0.1)) - 2.378) * (p_intSIHtFt - 4.5) + 4.5;
                dblSI = dblSI + 0.49;
                return dblSI;
            }

            /// <summary>
            /// SITE INDEX FOR DOUGLAS-FIR - Monserud
            /// Applying Height Growth and Site Index Curves for Inland DOUGLAS-FIR
            /// Research Paper:  INT-347  1985
            /// Derived from Kurt's PL/SQL
            /// Base age is 50 years
            /// site index curves validated by jsfried

            /// </summary>
            /// <param name="p_intSIDiaAge">Age of site tree (Ring count at breast height)</param>
            /// <param name="p_intSIHtFt">Diameter of site tree</param>
            /// <param name="p_strHabTypeCd">Habitat type code for condition</param>
            private double SI_DF2(int p_intSIDiaAge, int p_intSIHtFt, string p_strHabTypeCd)
            {
                double dblSI = 0;
                int intHabTypeCd = 0;
                // habTypeCd is stored as text in the cond table; Safely try to parse to an int
                bool isHabTypeAnInt = int.TryParse(p_strHabTypeCd, out intHabTypeCd);
                if (!isHabTypeAnInt)
                {
                    // if habTypeCd is not an int, set it to the middle value
                    intHabTypeCd = 500;
                }
                double dblC1 = 0;
                double dblC2 = 0;
                double dblC3 = 0;
                if (intHabTypeCd <= 400)
                {
                    dblC1 = 1.0;
                }
                else if (intHabTypeCd < 530)
                {
                    dblC2 = 1.0;
                }
                else if (intHabTypeCd >= 530)
                {
                    dblC3 = 1.0;
                }
                dblSI = (38.787-(2.805 * (Math.Log(p_intSIDiaAge) * Math.Log(p_intSIDiaAge))) +
                        (0.0216 * p_intSIDiaAge * Math.Log(p_intSIDiaAge)) +
                        (0.4948 * dblC1 + 0.4305 * dblC2 + 0.3964 * dblC3) * p_intSIHtFt +
                        (25.315 * dblC1 + 28.415 * dblC2 + 30.008 * dblC3) * p_intSIHtFt / p_intSIDiaAge) + 4.5;

                return dblSI;
            }

            /// <summary>
            /// This equation for Interior Western Red Cedar is from Hegyi and others
            /// Site Index Equations and curves for the Major Tree Species in B.C.
            /// -- B.C. Min of Forest Inventory Report 189(1) Nov 1979, Rev Sept 1981, p.6
            /// -- Site index at total tree age of 100 -- Equation #272
            /// ---Derived from Kurt's PL/SQL
            ///OUTPUT HAS NOT BEEN VALIDATED; METHOD NOT CURRENTLY IN USE
            ///<param name="p_intSIDiaAge"></param>
            ///<param name="p_intSIHtFt"></param>
            /// </summary>
            private double THPL03(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = 0;
                // PL/SQL:  HT/(1.3283*(1-EXP(-0.0174*AGE))**1.4711);
                dblSI = p_intSIHtFt/(1.3283 * Math.Pow(1 - Math.Exp(-0.0174*p_intSIDiaAge), 1.4711));
                return dblSI;
            }

            /// <summary>
            ///--- SITE INDEX RED ALDER - Harrington
		    ///    Height Growth and Site Index Curves for Red Alder
		    ///    Research Paper:  PNW-358 1986 p.5
            /// ---Derived from VBA source code by Don Vandendriese for FIA2FVS from RMRS
            ///OUTPUT HAS NOT BEEN VALIDATED; METHOD NOT CURRENTLY IN USE
            ///<param name="p_intSIDiaAge"></param>
            ///<param name="p_intSIHtFt"></param>
            /// </summary>
            private double SI_RA1(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = 0;
                double a = 54.185 - 4.61694 * p_intSIDiaAge + 0.11065 * Math.Pow(p_intSIDiaAge, 2) - 0.0007633 * Math.Pow(p_intSIDiaAge, 3);
                double b = 1.25934 - (0.012989) * p_intSIDiaAge + 3.522 * Math.Pow((1.0 / p_intSIDiaAge), 3);
                dblSI = a + b * p_intSIHtFt;
                dblSI = dblSI + 0.49;
                return dblSI;
            }


		}

        private IDictionary<String, String> LoadSiteIndexEquations(string strVariant)
        {
            //instantiate the dictionary so we can add equation records
            IDictionary<String, String> _dictSiteIdxEq = new Dictionary<String, String>();
            ado_data_access oAdo = new ado_data_access();
            //create env object so we can get the appDir
            env pEnv = new env();
            //open the project db file; db name is hard-coded
            oAdo.OpenConnection(oAdo.getMDBConnString(pEnv.strAppDir + "\\db\\ref_master.mdb", "", ""));
            string strSQL = "select * from site_index_equations where FVS_VARIANT = '" + strVariant + "'";
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, strSQL);
            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    if (oAdo.m_OleDbDataReader["FVS_VARIANT"] == System.DBNull.Value ||
                        oAdo.m_OleDbDataReader["FIA_SPCD"] == System.DBNull.Value ||
                        oAdo.m_OleDbDataReader["EQUATION"] == System.DBNull.Value)
                    {
                        //If either variant, spcd, or equation is null, we don't add because we can't use
                    }
                    else
                    {
                        string strFvsVariant = Convert.ToString(oAdo.m_OleDbDataReader["FVS_VARIANT"]);
                        string strFiaSpCd = Convert.ToString(oAdo.m_OleDbDataReader["FIA_SPCD"]);
                        string strRegion = SI_EMPTY;
                        if (oAdo.m_OleDbDataReader["REGION"] != System.DBNull.Value)
                        {
                            strRegion = Convert.ToString(oAdo.m_OleDbDataReader["REGION"]);
                        }
                        string strSlfSpcd = SI_EMPTY;
                        if (oAdo.m_OleDbDataReader["SLF_SPCD"] != System.DBNull.Value)
                        {
                            strSlfSpcd = Convert.ToString(oAdo.m_OleDbDataReader["SLF_SPCD"]);
                        }
                        string strEquation = Convert.ToString(oAdo.m_OleDbDataReader["EQUATION"]);
                        string strValue = strEquation + SI_DELIM + strSlfSpcd + SI_DELIM + strRegion;
                        _dictSiteIdxEq.Add(strFvsVariant + SI_DELIM + strFiaSpCd, strValue);
                    }
                }
            }
            // Always close the connection
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oAdo = null;
            return _dictSiteIdxEq;
        }
	}
	
	
			

	
}
