using System;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for ado_optimizer_tables.
	/// </summary>
	public class ado_optimizer_tables
	{
		public System.Data.DataSet m_dsOptimizerTables;
		public int m_intNumberOfOptimizerTablesLoaded;
		public string[] m_strOptimizerTables;
        public System.Data.OleDb.OleDbDataAdapter m_OleDbDataAdapter;

		public System.Data.OleDb.OleDbDataAdapter m_daPlot;
        public System.Data.OleDb.OleDbDataAdapter m_daCond;
		public System.Data.OleDb.OleDbDataAdapter m_daFFE;
		public System.Data.OleDb.OleDbDataAdapter m_daHarvestCosts;
		public System.Data.OleDb.OleDbDataAdapter m_daOwnerGroups;
		public System.Data.OleDb.OleDbDataAdapter m_daTreeClass;
		public System.Data.OleDb.OleDbDataAdapter m_daRX;
		public System.Data.OleDb.OleDbDataAdapter m_daTreeSpecies;
		public System.Data.OleDb.OleDbDataAdapter m_daTreeVol;
		public System.Data.OleDb.OleDbDataAdapter m_daTravelTimes;
		public System.Data.OleDb.OleDbDataAdapter m_daPSites;

		public System.Data.OleDb.OleDbCommand m_OleDbCommand;
		public System.Data.OleDb.OleDbConnection m_OleDbConnection;
		public System.Data.OleDb.OleDbConnection m_connPlot;
        public System.Data.OleDb.OleDbConnection m_connCond;
		public System.Data.OleDb.OleDbConnection m_connMasterLink;
		public System.Data.OleDb.OleDbConnection m_OleDbConnectionFFE;
		public System.Data.OleDb.OleDbConnection m_OleDbConnectionTravelTimes;
		public System.Data.OleDb.OleDbConnection m_OleDbConnectionProcessingSites;
		public System.Data.OleDb.OleDbConnection m_OleDbConnectionScenario;
		public System.Data.OleDb.OleDbDataReader m_OleDbDataReader;
		public System.Data.DataTable m_DataTable;
		public string m_strError;
		public int m_intError;
        public const int NUMBER_OF_OPTIMIZER_TABLES =  11;
		public string m_strRandomFileName="";
		public int m_intNumberOfOptimizerTables;

		public ado_optimizer_tables()
		{
			this.m_strOptimizerTables = new string[50];
			for (int x=0; x <= NUMBER_OF_OPTIMIZER_TABLES - 1; x++)
			{
				this.m_strOptimizerTables[x] = "";
			}
			this.m_intNumberOfOptimizerTablesLoaded = 0;

			this.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
			this.m_daPlot = new System.Data.OleDb.OleDbDataAdapter();
            this.m_daCond = new System.Data.OleDb.OleDbDataAdapter();
			this.m_daFFE = new System.Data.OleDb.OleDbDataAdapter();
			this.m_daHarvestCosts = new System.Data.OleDb.OleDbDataAdapter();
			this.m_daOwnerGroups = new System.Data.OleDb.OleDbDataAdapter();
			this.m_daTreeClass = new System.Data.OleDb.OleDbDataAdapter();
			this.m_daRX = new System.Data.OleDb.OleDbDataAdapter();
			this.m_daTreeSpecies = new System.Data.OleDb.OleDbDataAdapter();
		    this.m_daTreeVol = new System.Data.OleDb.OleDbDataAdapter();
			this.m_daTravelTimes = new System.Data.OleDb.OleDbDataAdapter();
			this.m_daPSites = new System.Data.OleDb.OleDbDataAdapter();

			this.m_connPlot = new System.Data.OleDb.OleDbConnection();
			this.m_connCond = new System.Data.OleDb.OleDbConnection();
			

			this.m_dsOptimizerTables = new DataSet();
			this.m_dsOptimizerTables.Clear();
			
			//
			// TODO: Add constructor logic here
			//
		}
        public void LoadDataSourceTables(string strScenarioMDB, string strScenarioId)
		{
			string strSQL="";
            string strFullPathMDB="";
			string strConn="";
			
              
			ado_data_access p_ado = new ado_data_access();

			            
			
			this.m_OleDbConnectionScenario = new System.Data.OleDb.OleDbConnection();


			strConn = p_ado.getMDBConnString(strScenarioMDB,"admin","");
			p_ado.OpenConnection(strConn, ref this.m_OleDbConnectionScenario);	
			if (p_ado.m_intError != 0)
			{
				p_ado = null;
				return;
			}
			strSQL = "SELECT table_name FROM scenario_datasource WHERE " + 
				" scenario_id = '" + strScenarioId + "';";
			p_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);

				  
			
			if (p_ado.m_intError == 0)
			{
				this.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
				this.m_dsOptimizerTables = new DataSet();
                this.m_OleDbCommand = new System.Data.OleDb.OleDbCommand();
				this.m_dsOptimizerTables.DataSetName = "LoadAllRecordsFromScenarioDataSource";
		        while (p_ado.m_OleDbDataReader.Read())
				 {
					if (p_ado.m_OleDbDataReader["table_name"] != System.DBNull.Value)
					{
						if (p_ado.m_OleDbDataReader["table_name"].ToString().Trim().Length > 0)
						{
							
							if (p_ado.m_OleDbDataReader["table_name"].ToString().Trim().ToUpper() == "PLOT" ||
								p_ado.m_OleDbDataReader["table_name"].ToString().Trim().ToUpper() == "COND")
							{
								strFullPathMDB = p_ado.m_OleDbDataReader["path"].ToString().Trim() + "\\" + 
									          p_ado.m_OleDbDataReader["file"].ToString().Trim();
								if (System.IO.File.Exists(strFullPathMDB) == true)
								{
                                    strConn = p_ado.getMDBConnString(strFullPathMDB,"admin","");
									p_ado.OpenConnection(strConn);
									if (p_ado.m_intError == 0)
									{
								        this.m_OleDbDataAdapter.SelectCommand.CommandText = "select * from " + 
											p_ado.m_OleDbDataReader["table_name"].ToString().Trim();
                                        this.m_OleDbDataAdapter.Fill(this.m_dsOptimizerTables,p_ado.m_OleDbDataReader["table_name"].ToString().Trim());

									}
									else 
									{
										this.m_intError=-1;
										p_ado.m_OleDbDataReader.Close();
										this.m_OleDbConnection.Close();
										p_ado.m_OleDbCommand = null;
										p_ado.m_OleDbDataReader = null;
										p_ado = null;
										this.m_dsOptimizerTables.Clear();
										this.m_dsOptimizerTables = null;
										this.m_OleDbDataAdapter.Dispose();
										this.m_OleDbDataAdapter=null;
										return;
									}
								}
								else
								{
									//MessageBox.Show(strFullPathMDB + " does not exist","Scenario Data Source", MessageBoxIcon.Error, MessageBoxButtons.OK);
									MessageBox.Show(strFullPathMDB + " does not exist","Scenario Data Source");
									this.m_intError=-1;
									p_ado.m_OleDbDataReader.Close();
									this.m_OleDbConnection.Close();
									p_ado.m_OleDbCommand = null;
									p_ado.m_OleDbDataReader = null;
									p_ado = null;

									this.m_dsOptimizerTables.Clear();
									this.m_dsOptimizerTables = null;
									this.m_OleDbDataAdapter.Dispose();
									this.m_OleDbDataAdapter = null;
									return;
								}

							}
						}
					}

				}
				
				p_ado.m_OleDbDataReader.Close();

			}
			p_ado.m_OleDbDataReader = null;
			p_ado.m_OleDbCommand = null;

			p_ado = null;
			this.m_OleDbConnectionScenario.Close();

		}

		public void LoadDataSourceTablesFromListBox(string strScenarioMDB, string strScenarioId,
			System.Windows.Forms.ListBox listBox1, string strDestinationLinkMDB)
		{
			string strSQL="";
			string strFullPathMDB="";
			string strConn="";
			
     		int x=0;
			int y=0;
			bool lLoaded=false;
			
			this.m_intError=0;

            this.m_intNumberOfOptimizerTables = this.getNumberOfOptimizerTables(strScenarioMDB, strScenarioId);
		
			//ado specific routines class
			ado_data_access p_ado = new ado_data_access();

			

			
            

			//connect to mdb file containing data sources
			this.m_OleDbConnectionScenario = new System.Data.OleDb.OleDbConnection();
			strConn = p_ado.getMDBConnString(strScenarioMDB,"admin","");
			
			p_ado.OpenConnection(strConn, ref this.m_OleDbConnectionScenario);	
			if (p_ado.m_intError != 0)
			{
				this.m_intError = p_ado.m_intError;
				p_ado = null;
				return;
			}

            //go through each of the items in the listbox
			for (y = 0; y <= listBox1.Items.Count - 1; y++)
			{
				lLoaded=false;
				//see if the listbox item is already loaded into a dataset table and the linked mdb table
				if (this.m_intNumberOfOptimizerTablesLoaded != 0)
				{   
					for (x=0;x<=this.m_intNumberOfOptimizerTables - 1;x++)
					{
						if (this.m_strOptimizerTables[x].Trim().Length > 0)
						{
							if (listBox1.Items[y].ToString().Trim().ToLower() == 
								this.m_strOptimizerTables[x].Trim().ToLower())
							{
                                lLoaded=true;
								break;
						}
					}
				}
			}
            if (lLoaded==false)
            {


                //query the MDB datasource table for Optimizer table names
			  strSQL = "SELECT path, file,table_name FROM scenario_datasource WHERE " + 
	        	    	" scenario_id = '" + strScenarioId + "' AND " + 
					    " table_name = '" + listBox1.Items[y] + "';";
			  p_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);
				if (p_ado.m_intError == 0)
				{
					//read the record
					while (p_ado.m_OleDbDataReader.Read())
					{
						//look to make sure we have the correct record
						if (listBox1.Items[y].ToString().Trim().ToUpper() 
							==
							p_ado.m_OleDbDataReader["table_name"].ToString().Trim().ToUpper())
						{
							strFullPathMDB = p_ado.m_OleDbDataReader["path"].ToString().Trim() + "\\" + 
								p_ado.m_OleDbDataReader["file"].ToString().Trim();
							if (System.IO.File.Exists(strFullPathMDB) == true)
							{
								//used to create a link to the table
								dao_data_access p_dao = new dao_data_access();

								//create a link to the table in an mdb file
								p_dao.CreateTableLink(strDestinationLinkMDB,listBox1.Items[y].ToString().Trim(),
									                   strFullPathMDB,listBox1.Items[y].ToString().Trim());
								p_dao = null;

								//connect to mdb file that will be used as the master table link file
								this.m_connMasterLink = new System.Data.OleDb.OleDbConnection();
								strConn=p_ado.getMDBConnString(strDestinationLinkMDB,"admin","");
								
								p_ado.OpenConnection(strConn, ref this.m_connMasterLink);
								strSQL = "SELECT * FROM " + p_ado.m_OleDbDataReader["table_name"];
								this.m_OleDbDataAdapter.SelectCommand = new System.Data.OleDb.OleDbCommand(strSQL,this.m_connMasterLink);
								try 
								{

     								this.m_OleDbDataAdapter.Fill(this.m_dsOptimizerTables,p_ado.m_OleDbDataReader["table_name"].ToString().Trim());
	    						}
		    					catch (Exception e)
			    				{
				    				MessageBox.Show(e.Message,"Table",MessageBoxButtons.OK,MessageBoxIcon.Error);
					    			this.m_intError=-1;
						    	}
								this.m_strOptimizerTables[this.m_intNumberOfOptimizerTablesLoaded] = p_ado.m_OleDbDataReader["table_name"].ToString().Trim();
								this.m_intNumberOfOptimizerTablesLoaded++;

								this.m_connMasterLink.Close();
                                this.m_connMasterLink = null;
								
							}
							else
							{
								
								MessageBox.Show(strFullPathMDB + " does not exist","Scenario Data Source");
								this.m_intError=-1;
								p_ado.m_OleDbDataReader.Close();
								this.m_OleDbConnection.Close();
								p_ado.m_OleDbCommand = null;
								p_ado.m_OleDbDataReader = null;
								p_ado = null;
								this.m_dsOptimizerTables.Clear();
								this.m_dsOptimizerTables = null;
								this.m_OleDbDataAdapter.Dispose();
								this.m_OleDbDataAdapter = null;
								return;
							}

						}
					}
					p_ado.m_OleDbDataReader.Close();
				}            
			}
				
            

			}
			p_ado.m_OleDbDataReader = null;
			p_ado.m_OleDbCommand = null;

			p_ado = null;
			this.m_OleDbConnectionScenario.Close();
			

		}
        public int getNumberOfOptimizerTables(string strScenarioMDB, string strScenarioId)
		{

			int intCount=0;
			System.Data.OleDb.OleDbConnection p_conn;
                        
		    ado_data_access p_ado = new ado_data_access();
			p_conn = new System.Data.OleDb.OleDbConnection();
			string strConn = p_ado.getMDBConnString(strScenarioMDB,"admin","");
			
			p_ado.OpenConnection(strConn, ref p_conn);	
			if (p_ado.m_intError != 0)
			{
				p_ado = null;
				return intCount;
			}
			string strSQL = "SELECT table_name FROM scenario_datasource WHERE " + 
				" scenario_id = '" + strScenarioId + "';";

			p_ado.SqlQueryReader(p_conn, strSQL);
			if (p_ado.m_intError == 0)
			{
				while (p_ado.m_OleDbDataReader.Read())
				{
					if (p_ado.m_OleDbDataReader["table_name"] != System.DBNull.Value)
					{
						if (p_ado.m_OleDbDataReader["table_name"].ToString().Trim().Length > 0)
						{
                             intCount++; 
						}
					}
				}
			}
			p_ado.m_OleDbDataReader.Close();
			p_ado.m_OleDbDataReader = null;
            p_conn.Close();
			p_conn=null;
            p_ado = null;
			return intCount;
		}
		public void CreateMDBAndCreateOptimizerTableDataSourceLinks(string strScenarioMDB, string strScenarioId,string strDestinationLinkDir)
		{
			string strSQL="";
			string strFullPathMDB="";
			string strConn="";

			//ado specific routines class
			ado_data_access p_ado = new ado_data_access();

			//connect to mdb file containing data sources
			this.m_OleDbConnectionScenario = new System.Data.OleDb.OleDbConnection();
			strConn=p_ado.getMDBConnString(strScenarioMDB,"admin","");
			
			p_ado.OpenConnection(strConn, ref this.m_OleDbConnectionScenario);	
			if (p_ado.m_intError != 0)
			{
				this.m_intError = p_ado.m_intError;
				p_ado = null;
				return;
			}

			//used to get the temporary random file name
			utils p_utils = new utils();


			//get temporary mdb file
			this.m_strRandomFileName = 
				p_utils.getRandomFile(strDestinationLinkDir,"accdb");
			p_utils = null;

			//used to create a link to the table
			dao_data_access p_dao = new dao_data_access();

			//create a temporary mdb that will contain all 
            //the links to the Optimizer tables
			p_dao.CreateMDB(this.m_strRandomFileName);


            //query the MDB datasource table for Optimizer table names and location of the table
			
			strSQL = "SELECT path, file,table_name FROM scenario_datasource WHERE " + 
				" scenario_id = '" + strScenarioId + "';";
			p_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);
			if (p_ado.m_intError == 0)
			{
				//read the record
				while (p_ado.m_OleDbDataReader.Read())
				{
					//look to make sure we have the correct record
					strFullPathMDB = p_ado.m_OleDbDataReader["path"].ToString().Trim() + "\\" + 
						p_ado.m_OleDbDataReader["file"].ToString().Trim();
					if (System.IO.File.Exists(strFullPathMDB) == true)
					{
						if (p_ado.m_OleDbDataReader["table_name"].ToString().Trim().ToUpper() == "PLOT" ||
							p_ado.m_OleDbDataReader["table_name"].ToString().Trim().ToUpper() == "COND" ||
							p_ado.m_OleDbDataReader["table_name"].ToString().Trim().ToUpper() == "FFE")

						{						
							//create a link to the table in an mdb file
							p_dao.CreateTableLink(this.m_strRandomFileName,p_ado.m_OleDbDataReader["table_name"].ToString().Trim(),
								strFullPathMDB,p_ado.m_OleDbDataReader["table_name"].ToString().Trim());
							
						}
					}
				}
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbDataReader=null;
				
			}
			this.m_OleDbConnectionScenario.Close();
			this.m_OleDbConnectionScenario=null;
			p_ado=null;
			p_dao = null;

		}
		
	
	}


}
