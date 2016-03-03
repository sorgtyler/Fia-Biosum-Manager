using System;
using System.Windows.Forms;
using System.Data;
using System.IO;
using FIA_Biosum_Manager;
namespace FIA_Biosum_Data_Converter
{
	
	/// <summary>
	/// Summary description for ado_data_access.
	/// </summary>
	public class convert_data
	{
		public System.Data.DataSet m_DataSet;
		public System.Data.OleDb.OleDbDataAdapter m_OleDbDataAdapter;
		public System.Data.OleDb.OleDbCommand m_OleDbCommand;
		public System.Data.OleDb.OleDbConnection m_OleDbSourceConn;
		public System.Data.OleDb.OleDbConnection m_OleDbTargetConn;
		private System.Data.OleDb.OleDbConnection m_TempMDBFileConn;

		private System.Data.OleDb.OleDbDataReader m_PnwIdbDataReader;
		public System.Data.DataTable m_DataTable;
		public string m_strError;
		public int m_intError;
		public string strTempMDBFile;
		private	FIA_Biosum_Manager.frmTherm m_frmTherm;
		private ado_data_access m_ado;
		private dao_data_access m_dao;
		private System.IO.FileStream m_txtFileStream;
		private System.IO.StreamWriter m_txtStreamWriter;
        
		public convert_data()
		{
			this.m_ado = new ado_data_access();
			this.m_dao = new dao_data_access();
            this.m_TempMDBFileConn = new System.Data.OleDb.OleDbConnection();
			//make a new FileStream object, ready for read and write steps. Dim fs As 
		    this.m_txtFileStream = new FileStream("c:\\fia_biosum\\orca_converted_data\\convert_data_log.txt", FileMode.Create, 
		                                 FileAccess.Write);
			this.m_txtStreamWriter = new StreamWriter(this.m_txtFileStream);
			this.m_txtStreamWriter.Write("Convert Oregon And California Data " + System.DateTime.Now.ToString() + "\n");
			this.m_txtStreamWriter.Write("----------------------------------------------------------\n");


		}
	  ~convert_data()
		{
		  try
		  {
			  this.m_txtStreamWriter.Close();
			  this.m_txtFileStream.Close();
		  }
		  catch
		  {
		  }

		}
		public void convert_all()
		{
			this.m_intError=0;

			m_frmTherm = new FIA_Biosum_Manager.frmTherm();
			m_frmTherm.btnCancel.Visible=true;
			m_frmTherm.Show();
			m_frmTherm.Focus();
			m_frmTherm.Text = "Converting Oregon And California Previous Study Data";
			m_frmTherm.Refresh();
			m_frmTherm.progressBar1.Minimum = 1;
			this.m_frmTherm.btnCancel.Click += new System.EventHandler(this.ThermCancel);


			string strIdbMDBFile = "c:\\fia_biosum\\orca_converted_data\\db\\idb version 2.mdb";
		    string strSourceMDBFile = "c:\\fia_biosum\\orca_converted_data\\db\\orca4_core.mdb";
		    string strTargetMDBFile = "c:\\fia_biosum\\orca_converted_data\\db\\master.mdb";
			string strScenarioMDBFile = "c:\\fia_biosum\\orca_converted_data\\core\\db\\scenario.mdb";
			string strHaulCostMDBFile = "c:\\fia_biosum\\orca_converted_data\\db\\orca4_haul.mdb";
			string strTargetTravelTimesMDBFile = "c:\\fia_biosum\\orca_converted_data\\gis\\db\\gis_travel_times.mdb";
			string strIdbToNimsMDBFile = "c:\\fia_biosum\\orca_converted_data\\db\\idb to nims conversion.mdb";
			this.m_txtStreamWriter.Write ("Using These MDB Files:\n");
			this.m_txtStreamWriter.Write ("IDB Source:" + strIdbMDBFile + "\n Previous Study Source:" + strSourceMDBFile + " \n Target Source:" + strTargetMDBFile + "\n" + "Haul Cost Source:" + strHaulCostMDBFile + "\n" + "IDB To Nims Source:" + strIdbToNimsMDBFile);
			utils p_utils = new utils();
			env p_env = new env();

			this.strTempMDBFile = p_utils.getRandomFile(p_env.strTempDir,"mdb");
			this.m_txtStreamWriter.Write("Creating MDB file and table links to " + this.strTempMDBFile + "\n");
			m_dao.CreateMDB(this.strTempMDBFile);
			m_dao.CreateTableLink(this.strTempMDBFile,"pnw_idb_plot",strIdbMDBFile,"PLOT");
			m_dao.CreateTableLink(this.strTempMDBFile,"pnw_idb_cond",strIdbMDBFile,"COND");
			m_dao.CreateTableLink(this.strTempMDBFile,"target_plot",strTargetMDBFile,"plot");
			m_dao.CreateTableLink(this.strTempMDBFile,"inventories",strTargetMDBFile,"inventories");
			m_dao.CreateTableLink(this.strTempMDBFile,"source_plot", strSourceMDBFile,"all_plots");
			m_dao.CreateTableLink(this.strTempMDBFile,"source_harvest_costs", strSourceMDBFile,"harvest_costs");
            m_dao.CreateTableLink(this.strTempMDBFile,"source_cond", strSourceMDBFile,"all_conditions");
			m_dao.CreateTableLink(this.strTempMDBFile,"source_ffe", strSourceMDBFile,"ffe");
            m_dao.CreateTableLink(this.strTempMDBFile,"target_cond",strTargetMDBFile,"cond");
			m_dao.CreateTableLink(this.strTempMDBFile,"target_ffe",strTargetMDBFile,"ffe");
			m_dao.CreateTableLink(this.strTempMDBFile,"source_tree_diam_groups", strSourceMDBFile,"ref_diameter_classes");
			m_dao.CreateTableLink(this.strTempMDBFile,"target_tree_diam_groups",strTargetMDBFile,"tree_diam_groups");
			m_dao.CreateTableLink(this.strTempMDBFile,"source_rx", strSourceMDBFile,"ref_rx_description");
			m_dao.CreateTableLink(this.strTempMDBFile,"target_rx",strTargetMDBFile,"rx");
			m_dao.CreateTableLink(this.strTempMDBFile,"target_rx_intensity",strScenarioMDBFile,"scenario_rx_intensity");
			m_dao.CreateTableLink(this.strTempMDBFile,"source_tree_species_groups", strSourceMDBFile,"ref_spp_grp");
			m_dao.CreateTableLink(this.strTempMDBFile,"target_tree_species_groups",strTargetMDBFile,"tree_species_groups");
			m_dao.CreateTableLink(this.strTempMDBFile,"target_harvest_costs",strTargetMDBFile,"harvest_costs");
			m_dao.CreateTableLink(this.strTempMDBFile,"source_tree_vol_val_by_species_diam_groups", strSourceMDBFile,"vol_val_by_spp_diam");
			m_dao.CreateTableLink(this.strTempMDBFile,"target_tree_vol_val_by_species_diam_groups",strTargetMDBFile,"tree_vol_val_by_species_diam_groups");
			m_dao.CreateTableLink(this.strTempMDBFile,"target_travel_times_plot",strTargetTravelTimesMDBFile,"plot");
			m_dao.CreateTableLink(this.strTempMDBFile,"target_psite",strTargetTravelTimesMDBFile,"processing_site");
			m_dao.CreateTableLink(this.strTempMDBFile,"source_psite",strSourceMDBFile,"psite_lu");
      m_dao.CreateTableLink(this.strTempMDBFile,"source_haul_cost", strHaulCostMDBFile,"haul");
      m_dao.CreateTableLink(this.strTempMDBFile,"source_nims_tree",strIdbToNimsMDBFile,"NIMS_TREE");
      m_dao.CreateTableLink(this.strTempMDBFile,"target_tree",strTargetMDBFile,"tree");

            m_ado.OpenConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.strTempMDBFile + ";User Id=admin;Password=;", ref this.m_TempMDBFileConn);
			if (m_ado.m_intError==0) this.convert_plot();
		    if (m_frmTherm.AbortProcess==false && m_ado.m_intError==0 && this.m_intError==0) this.convert_cond();
			if (m_frmTherm.AbortProcess==false && m_ado.m_intError==0 && this.m_intError==0) this.convert_ffe();
		    if (m_frmTherm.AbortProcess==false && m_ado.m_intError==0 && this.m_intError==0) this.convert_treediamgroups();
		    if (m_frmTherm.AbortProcess==false && m_ado.m_intError==0 && this.m_intError==0) this.convert_rx();
			if (m_frmTherm.AbortProcess==false && m_ado.m_intError==0 && this.m_intError==0) this.convert_treespeciesgroups();
			if (m_frmTherm.AbortProcess==false && m_ado.m_intError==0 && this.m_intError==0) this.convert_harvestcosts();
			if (m_frmTherm.AbortProcess==false && m_ado.m_intError==0 && this.m_intError==0) this.convert_treevolval();
			if (m_frmTherm.AbortProcess==false && m_ado.m_intError==0 && this.m_intError==0) this.convert_haulcost();
			if (m_frmTherm.AbortProcess==false && m_ado.m_intError==0 && this.m_intError==0) this.create_tree();
			this.m_frmTherm.Close();
			try
			{
				this.m_txtStreamWriter.Close();
				this.m_txtFileStream.Close();
			}
			catch
			{
			}

			
		}
		private void convert_plot()
		{
			string str;
			System.Data.OleDb.OleDbDataAdapter p_da;
			System.Data.DataSet p_ds;
			System.Data.OleDb.OleDbCommand p_OleDbCommand;
            p_da = new System.Data.OleDb.OleDbDataAdapter();
			p_ds = new DataSet("tempds");
			this.m_intError=0;
			this.m_strError = "";
            string strSQL="";
			int intCurrRec=0;
			string strbiosumid, strplotid, strstatecd, strcountycd;
            string strplot,strmeasyear, strmeasmon, strmeasday;
            string strelevft, strfvsvariant,strhalfstate,strsubplotcountplot;
			string strgisyarddist, strgisdismovedtoroad,strgisprotectedarea;
			string strmovedfromprotected,strgisroadless,strgissteeporfar;
			string strgisaccessible,strnumcond,stronecond,strlat,strlon;
			string strDataSourceName, strforestblm,strunitcd;
			string strInvId="";
			string strDbFields="";
			string strValues="";
			string strDbTravelTimesFields="";
			string strDbTravelTimesValues="";
			
            this.m_txtStreamWriter.Write("Converting PLOT Data\n");
			this.m_txtStreamWriter.Write("--------------------\n");

			try
			{
				strSQL = "DELETE * FROM TARGET_PLOT";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, strSQL);
				strSQL = "SELECT * FROM SOURCE_PLOT";
				m_frmTherm.AbortProcess = false;
				m_frmTherm.lblMsg.Text = "Converting Plot Data";
				m_frmTherm.lblMsg.Visible = true;
				this.m_frmTherm.progressBar1.Maximum = Convert.ToInt32(this.m_ado.getRecordCount("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.strTempMDBFile + ";User Id=admin;Password=;","select count(*) from source_plot","source_plot"));
				this.m_txtStreamWriter.Write("source plot record count=" + Convert.ToString(this.m_frmTherm.progressBar1.Maximum) + "\n");
				p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
				p_OleDbCommand.CommandText = strSQL;
				p_da.SelectCommand = p_OleDbCommand;
				p_da.Fill(p_ds, "source_plot");
			}
			catch (Exception caught)
			{
				MessageBox.Show("Abort Process: Error in plot conversion with error " + caught.Message);
				this.m_intError=-1;
				p_OleDbCommand = null;
				p_ds = null;
				p_da = null;
		        return;
			}
			if (this.m_intError==0)
			{
				for (intCurrRec = 0; intCurrRec <= p_ds.Tables["source_plot"].Rows.Count-1;intCurrRec++)
				{
					strbiosumid="";
					strplotid=""; 
					strstatecd="";
					strcountycd="";
					strplot="";
					strmeasyear="";
					strmeasmon="";
					strmeasday=""; 
					strelevft="";
					strfvsvariant="";
					strhalfstate="";
					strsubplotcountplot="";
					strgisyarddist="";
					strgisyarddist="";
					strgisprotectedarea="";
					strmovedfromprotected="";
					strgisroadless = "";
					strgissteeporfar = "";
					strgisaccessible ="";
					strnumcond = "";
					stronecond = "";
					strlat="";
					strlon="";
					strunitcd="";
					this.m_frmTherm.Increment(intCurrRec + 1);
					
					strplotid = Convert.ToString(p_ds.Tables["source_plot"].Rows[intCurrRec]["idb_plot"]); //     this.m_ado.m_OleDbDataReader["idb_plot"]);
					strstatecd = Convert.ToString(p_ds.Tables["source_plot"].Rows[intCurrRec]["state"]); //this.m_ado.m_OleDbDataReader["state"]);
					strcountycd = Convert.ToString(p_ds.Tables["source_plot"].Rows[intCurrRec]["cty"]); //this.m_ado.m_OleDbDataReader["cty"]);
					strplot = Convert.ToString(p_ds.Tables["source_plot"].Rows[intCurrRec]["plot"]); //Convert.ToString(this.m_ado.m_OleDbDataReader["plot"]);
					strlat = Convert.ToString(p_ds.Tables["source_plot"].Rows[intCurrRec]["z10_x"]); //Convert.ToString(this.m_ado.m_OleDbDataReader["z10_x"]);
					strlon = Convert.ToString(p_ds.Tables["source_plot"].Rows[intCurrRec]["z10_y"]); //Convert.ToString(this.m_ado.m_OleDbDataReader["z10_y"]);
					stronecond = Convert.ToString(p_ds.Tables["source_plot"].Rows[intCurrRec]["onecond"]); //Convert.ToString(this.m_ado.m_OleDbDataReader["onecond"]);
					strnumcond = Convert.ToString(p_ds.Tables["source_plot"].Rows[intCurrRec]["num_conditions"]); //Convert.ToString(this.m_ado.m_OleDbDataReader["num_conditions"]);
					strgisaccessible = Convert.ToString(p_ds.Tables["source_plot"].Rows[intCurrRec]["accessible"]); //Convert.ToString(this.m_ado.m_OleDbDataReader["accessible"]);
					if (strgisaccessible.IndexOf("T",0) < 0)
					{
						strgisaccessible="N";
					}
					else
					{
						strgisaccessible="Y";
					}
					strgissteeporfar = Convert.ToString(p_ds.Tables["source_plot"].Rows[intCurrRec]["steep_far"]); //Convert.ToString(this.m_ado.m_OleDbDataReader["steep_far"]);
					if (strgissteeporfar.IndexOf("T",0) < 0)
					{
						strgissteeporfar="N";
					}
					else
					{
						strgissteeporfar="Y";
					}
					strgisprotectedarea = Convert.ToString(p_ds.Tables["source_plot"].Rows[intCurrRec]["protected_area"]); //Convert.ToString(this.m_ado.m_OleDbDataReader["protected_area"]);
					if (strgisprotectedarea.IndexOf("T",0) < 0)
					{
						strgisprotectedarea="N";
					}
					else
					{
						strgisprotectedarea="Y";
					}

					strgisroadless = Convert.ToString(p_ds.Tables["source_plot"].Rows[intCurrRec]["rdls"]); //Convert.ToString(this.m_ado.m_OleDbDataReader["rdls"]);
					if (strgisroadless.Trim() == "0")
					{
						strgisroadless="N";
					}
					else
					{
						strgisroadless="Y";
					}
					stronecond = Convert.ToString(p_ds.Tables["source_plot"].Rows[intCurrRec]["onecond"]); //Convert.ToString(this.m_ado.m_OleDbDataReader["rdls"]);
					if (stronecond.IndexOf("T",0) < 0)
					{
						stronecond="N";
					}
					else
					{
						stronecond="Y";
					}

					strmovedfromprotected = Convert.ToString(p_ds.Tables["source_plot"].Rows[intCurrRec]["in_out_code"]); //Convert.ToString(this.m_ado.m_OleDbDataReader["in_out_code"]);
					if (strmovedfromprotected.Trim()=="0")
					{
						strmovedfromprotected = "N";
					}
					else
					{
						strmovedfromprotected="Y";
					}

					strSQL = "select assessment_date,elev_ft,fvs_variant," + 
						"forest_or_blm_district,data_source_name," +
						"subplot_count_plot,half_state,survey_unit_fia from pnw_idb_plot where " + 
						"plot_id = " + strplotid + ";";

					p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
					p_OleDbCommand.CommandText = strSQL;
					p_da.SelectCommand = p_OleDbCommand;
					p_da.Fill(p_ds, "pnw_idb_plot");
					if (p_ds.Tables["pnw_idb_plot"].Rows.Count == 1)
					{
						str = Convert.ToString(p_ds.Tables["pnw_idb_plot"].Rows[0]["assessment_date"]);
						if (str.Trim().Length > 0)
						{
							strmeasyear = str.Substring(str.LastIndexOf("/")+1,4);
							strmeasmon = str.Substring(0,str.IndexOf("/",0));
							strmeasday = str.Substring(str.IndexOf("/",0) + 1, str.Trim().Length - strmeasyear.Trim().Length - strmeasmon.Trim().Length - 2);
						}
						else
						{
							strmeasyear="";
							strmeasmon="";
							strmeasday="";
						}
						strelevft =  Convert.ToString(p_ds.Tables["pnw_idb_plot"].Rows[0]["elev_ft"]);
						strDataSourceName = Convert.ToString(p_ds.Tables["pnw_idb_plot"].Rows[0]["data_source_name"]);
						strfvsvariant = Convert.ToString(p_ds.Tables["pnw_idb_plot"].Rows[0]["fvs_variant"]);
						strhalfstate = Convert.ToString(p_ds.Tables["pnw_idb_plot"].Rows[0]["half_state"]);
						strsubplotcountplot = Convert.ToString(p_ds.Tables["pnw_idb_plot"].Rows[0]["subplot_count_plot"]);
						strforestblm = Convert.ToString(p_ds.Tables["pnw_idb_plot"].Rows[0]["forest_or_blm_district"]);
						strunitcd = Convert.ToString(p_ds.Tables["pnw_idb_plot"].Rows[0]["survey_unit_fia"]);

						strSQL = "select inv_id from inventories where " + 
							"trim(ucase(inv_id_def)) = '" + strDataSourceName.Trim().ToUpper() + "';";
						p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
						p_OleDbCommand.CommandText = strSQL;
						p_da.SelectCommand = p_OleDbCommand;
						p_da.Fill(p_ds, "inventories");
						if (p_ds.Tables["inventories"].Rows.Count == 1)
						{
							strInvId = Convert.ToString(p_ds.Tables["inventories"].Rows[0]["inv_id"]);
							if (str.Trim().Length == 0)
							{
								if (strmeasyear.Trim().Length == 4)
								{
									strInvId = strmeasyear;
								}
								else 
								{
									strInvId = "9999";
								}
							}

						}
						else
						{
							this.m_txtStreamWriter.Write("Data Error: could not find IDB data source name in inventories table for plot_id " + strplotid + "\n");
							if (strmeasyear.Trim().Length == 4)
							{
								strInvId = strmeasyear;
							}
							else 
							{
								strInvId = "9999";
							}

							//MessageBox.Show("!!!Could Not Find Plot Id = " + strplotid + " in idb version 2 plot table!!!");
						}
						p_ds.Tables["inventories"].Clear();


					}
					else
					{
						strmeasyear = "";
						strmeasyear = "";
						strmeasmon =  "";
						strmeasmon =  "";
						strmeasday =  "";
						strmeasday =  "";
						strelevft =  "";
						strDataSourceName = "";
						strfvsvariant = "";
						strhalfstate = "";
						strforestblm="";
						strsubplotcountplot = "";
						strunitcd="";
						this.m_txtStreamWriter.Write("Data Error: Could Not Find Plot Id = " + strplotid + " in idb version 2 plot table!!! \n");
					}
					p_ds.Tables["pnw_idb_plot"].Clear();
                     

					strSQL = "select yarding_distance from source_harvest_costs where " + 
						"idb_id = " + strplotid + ";";

					p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
					p_OleDbCommand.CommandText = strSQL;
					p_da.SelectCommand = p_OleDbCommand;
					p_da.Fill(p_ds, "source_harvest_costs");
					if (p_ds.Tables["source_harvest_costs"].Rows.Count > 0)
					{
						strgisyarddist = Convert.ToString(p_ds.Tables["source_harvest_costs"].Rows[0]["yarding_distance"]);
						if (strgisyarddist.Trim().Length == 0) strgisyarddist = "";
					}
					else
					{
						strgisyarddist="";
						this.m_txtStreamWriter.Write("Data Error: Could Not Find Plot Id = " + strplotid + " in harvest_costs table!!! \n");
					}
					p_ds.Tables["source_harvest_costs"].Clear();

					if (strInvId.Trim().Length == 0) strInvId = "9999";
					//format strbiosumid
					strbiosumid = "2" + strInvId;
					//state
					switch (strstatecd.Trim().Length)
					{
						case 0:
							strbiosumid = strbiosumid + "99";
							break;
						case 1:
							strbiosumid =  strbiosumid + "0" + strstatecd.Trim();
							break;
						default:
							strbiosumid = strbiosumid + strstatecd.Trim();
							break;
					}
					//cycle
					strbiosumid = strbiosumid + "00";
					//subcycle
					strbiosumid = strbiosumid + "00";
					//countycode
					switch (strcountycd.Trim().Length)
					{
						case 0:
							strbiosumid = strbiosumid + "999";
							break;
						case 1:
							strbiosumid = strbiosumid + "00" + strcountycd.Trim();
							break;
						case 2:
							strbiosumid = strbiosumid + "0" + strcountycd.Trim();
							break;
						default:
							strbiosumid = strbiosumid + strcountycd.Trim();
							break;
					}
					//plot
					switch (strplot.Trim().Length)
					{
						case 0:
							strbiosumid = strbiosumid + "9999999";
							break;
						case 1:
							strbiosumid = strbiosumid + "000000" + strplot.Trim();
							break;
						case 2:
							strbiosumid = strbiosumid + "00000" + strplot.Trim();
							break;
						case 3:
							strbiosumid = strbiosumid + "0000" + strplot.Trim();
							break;
						case 4:
							strbiosumid = strbiosumid + "000" + strplot.Trim();
							break;
						case 5:
							strbiosumid = strbiosumid + "00" + strplot.Trim();
							break;
						case 6:
							strbiosumid = strbiosumid + "0" + strplot.Trim();
							break;
						default:
							strbiosumid = strbiosumid + strplot.Trim();
							break;
					}

					//forest or blm district
					switch (strforestblm.Trim().Length)
					{
						case 0:
							strbiosumid = strbiosumid + "999";
							break;
						case 1:
							strbiosumid = strbiosumid + "00" + strforestblm.Trim();
							break;
						case 2:
							strbiosumid = strbiosumid + "0" + strforestblm.Trim();
							break;
						default:
							strbiosumid = strbiosumid + strforestblm.Trim();
							break;
					}


					strDbFields = "INSERT INTO target_plot (biosum_plot_id,idb_plot_id";
					strValues="VALUES ('" + strbiosumid + "'," + strplotid ;


					if (strstatecd.Trim().Length > 0)
					{
						strDbFields += ",statecd";
						strValues += "," + strstatecd;
						strDbTravelTimesFields += ",statecd";
						strDbTravelTimesValues += "," + strstatecd;

					}
					if (strcountycd.Trim().Length > 0)
					{
						strDbFields += ",countycd";
						strValues += "," + strcountycd;
						strDbTravelTimesFields += ",countycd";
						strDbTravelTimesValues += "," + strcountycd;

					}
					if (strplot.Trim().Length > 0)
					{
						strDbFields += ",plot";
						strValues += "," + strplot;
						strDbTravelTimesFields += ",plot";
						strDbTravelTimesValues += "," + strplot;

					}
					if (strmeasyear.Trim().Length > 0)
					{
						strDbFields += ",measyear";
						strValues += "," + strmeasyear;
						strDbTravelTimesFields += ",measyear";
						strDbTravelTimesValues += "," + strmeasyear;

					}
                    
					if (strmeasmon.Trim().Length > 0)
					{
						strDbFields += ",measmon";
						strValues += "," + strmeasmon;
					}
					if (strmeasday.Trim().Length > 0)
					{
						strDbFields += ",measday";
						strValues += "," + strmeasday;
					}
					if (strelevft.Trim().Length > 0)
					{
						strDbFields += ",elev";
						strValues += "," + strelevft;
						strDbTravelTimesFields += ",elevft";
						strDbTravelTimesValues += "," + strelevft;

					}
					if (strfvsvariant.Trim().Length > 0)
					{
						strDbFields += ",fvs_variant";
						strValues += ",'" + strfvsvariant + "'";
					}
					if (strhalfstate.Trim().Length > 0)
					{
						strDbFields += ",half_state";
						strValues += ",'" + strhalfstate + "'";
					}

					if (strsubplotcountplot.Trim().Length > 0)
					{
						strDbFields += ",subplot_count_plot";
						strValues += "," + strsubplotcountplot;
					}
					if (strgisyarddist.Trim().Length > 0)
					{
						strDbFields += ",gis_yard_dist";
						strValues += "," + strgisyarddist;
						strDbTravelTimesFields += ",gis_yard_dist";
						strDbTravelTimesValues += "," + strgisyarddist;

					}
					if (strgisyarddist.Trim().Length > 0)
					{
						strDbFields += ",gis_dist_moved_to_road";
						strValues += "," + strgisyarddist;
						strDbTravelTimesFields += ",gis_dist_moved_to_road";
						strDbTravelTimesValues += "," + strgisyarddist;

					}
					if (strgisprotectedarea.Trim().Length > 0)
					{
						strDbFields += ",gis_protected_area_yn";
						strValues += ",'" + strgisprotectedarea + "'";
						strDbTravelTimesFields += ",gis_protected_area_yn";
						strDbTravelTimesValues += ",'" + strgisprotectedarea + "'";

					}
					else
					{
						strDbFields += ",gis_protected_area_yn";
						strValues += ",'N'";
						strDbTravelTimesFields += ",gis_protected_area_yn";
						strDbTravelTimesValues += ",'N'";

					}
					if (strmovedfromprotected.Trim().Length > 0)
					{
						strDbFields += ",gis_moved_from_protected_to_unprotected_yn";
						strValues += ",'" + strmovedfromprotected + "'";
						strDbTravelTimesFields += ",gis_moved_from_protected_to_unprotected_yn";
						strDbTravelTimesValues += ",'" + strmovedfromprotected + "'";

					}
					else
					{
						strDbFields += ",gis_moved_from_protected_to_unprotected_yn";
						strValues += ",'N'";
						strDbTravelTimesFields += ",gis_moved_from_protected_to_unprotected_yn";
						strDbTravelTimesValues += ",'N'";

					}
					if (strgisroadless.Trim().Length > 0)
					{
						strDbFields += ",gis_roadless_yn";
						strValues += ",'" + strgisroadless + "'";
						strDbTravelTimesFields += ",gis_roadless_yn";
						strDbTravelTimesValues += ",'" + strgisroadless + "'";

					}
					else
					{
						strDbFields += ",gis_roadless_yn";
						strValues += ",'N'";
						strDbTravelTimesFields += ",gis_roadless_yn";
						strDbTravelTimesValues += ",'N'";

					}
					if (strgissteeporfar.Trim().Length > 0)
					{
						strDbFields += ",gis_steep_or_far_yn";
						strValues += ",'" + strgissteeporfar + "'";
						strDbTravelTimesFields += ",gis_steep_or_far_yn";
						strDbTravelTimesValues += ",'" + strgissteeporfar + "'";

					}
					else
					{
						strDbFields += ",gis_steep_or_far_yn";
						strValues += ",'N'";
						strDbTravelTimesFields += ",gis_steep_or_far_yn";
						strDbTravelTimesValues += ",'N'";

					}
					if (strgisaccessible.Trim().Length > 0)
					{
						strDbFields += ",gis_accessible_yn";
						strValues += ",'" + strgisaccessible + "'";
						strDbTravelTimesFields += ",gis_accessible_yn";
						strDbTravelTimesValues += ",'" + strgisaccessible + "'";

					}
					else
					{
						strDbFields += ",gis_accessible_yn";
						strValues += ",'N'";
						strDbTravelTimesFields += ",gis_accessible_yn";
						strDbTravelTimesValues += ",'N'";

					}
					if (strnumcond.Trim().Length > 0)
					{
						strDbFields += ",num_cond";
						strValues += "," + strnumcond;
					}
					if (stronecond.Trim().Length > 0)
					{
						strDbFields += ",one_cond_yn";
						strValues += ",'" + stronecond + "'";
					}
					else
					{
						strDbFields += ",one_cond_yn";
						strValues += ",'N'";
					}

					if (strlat.Trim().Length > 0)
					{
						strDbFields += ",lat";
						strValues += "," + strlat;
						strDbTravelTimesFields += ",lat";
						strDbTravelTimesValues += "," + strlat;

					}
					if (strlon.Trim().Length > 0)
					{
						strDbFields += ",lon";
						strValues += "," + strlon;
						strDbTravelTimesFields += ",lon";
						strDbTravelTimesValues += "," + strlon;

					}
					if (strunitcd.Trim().Length > 0)
					{
						strDbFields += ",unitcd";
						strValues += "," + strunitcd;
						strDbTravelTimesFields += ",unitcd";
						strDbTravelTimesValues += "," + strunitcd;

					}
					strDbFields += ",biosum_status_cd,gis_status_id)";
					strValues += ",1,1)";
					strDbTravelTimesFields += ",gis_status_id)";
					strDbTravelTimesValues += ",1)";

					strSQL = strDbFields + strValues;

					this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strSQL);
					if (this.m_ado.m_intError < 0)
					{
						this.m_txtStreamWriter.Write("Data Error: insert sql error for sql statement " + strSQL + "\n");
					}

					System.Windows.Forms.Application.DoEvents();
					if (this.m_frmTherm.AbortProcess == true) break;
				}
			}
		    p_OleDbCommand = null;
		    p_ds.Clear();
		    p_da.Dispose();
		    p_ds = null;
		    p_da = null;
		    this.m_txtStreamWriter.Write("Plot Records Processed: " + intCurrRec.ToString() + "/n");
		}
		private void convert_cond()
		{
			System.Data.OleDb.OleDbDataAdapter p_da;
			System.Data.DataSet p_ds;
			System.Data.OleDb.OleDbCommand p_OleDbCommand;
			p_da = new System.Data.OleDb.OleDbDataAdapter();
			p_ds = new DataSet("tempds");
			this.m_intError=0;
			this.m_strError = "";
			string strSQL="";
			int intCurrRec=0;

			string strbiosumcondid,strbiosumid, strplotid, strcondid,strcondnum;
			string strcondprop, strlandclcd,strfortypcd;
			string strglc, strowncd, strowngrpcd, strreservcd;
			string strsiteclcd,strsicond, strsisp;
			string strslope,straspect,strstdage,strstdszcd;
			string strhabtypcd1, stradforcd, strqmdtotcm;
			string stracres, strunitcd, strvollocgrp;
			string strtpa, strvolacgrsstemttlft3;
			string strvolacgrsft3,strvolacgrsstemsawlogft3;
			string strprenotcalc_yn,strpretotflamesev;
			string strpretotflamemod, strpretorchidx, strprecrownidx;
			string strprecanopyht, strprecanopydensity, strpremortbasev;
			string strpremortbamod, strpremortvolsev, strpremortvolmod;
			string strqmdhwdcm, strqmdswdcm, strbaft2ac;
			string strprefiretypesev,strprefiretypemod;
		    string strsdi,strccf,strtopht;

			string strDbFields="";
			string strValues="";
			
			this.m_intError=0;
			this.m_txtStreamWriter.Write("\n\nConverting COND Data\n");
			this.m_txtStreamWriter.Write("--------------------\n");

			try
			{
				strSQL = "DELETE * FROM TARGET_COND";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, strSQL);
				strSQL = "SELECT * FROM SOURCE_COND";
				m_frmTherm.AbortProcess = false;
				m_frmTherm.lblMsg.Text = "Converting Cond Data";
				m_frmTherm.lblMsg.Visible = true;
				m_frmTherm.lblMsg.Refresh();
				this.m_frmTherm.progressBar1.Maximum = Convert.ToInt32(this.m_ado.getRecordCount("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.strTempMDBFile + ";User Id=admin;Password=;","select count(*) from source_cond","source_cond"));
				this.m_txtStreamWriter.Write("source cond record count=" + Convert.ToString(this.m_frmTherm.progressBar1.Maximum) + "\n");
				p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
				p_OleDbCommand.CommandText = strSQL;
				p_da.SelectCommand = p_OleDbCommand;
				p_da.Fill(p_ds, "source_cond");
			}
			catch (Exception caught)
			{
				MessageBox.Show("Abort Process: Error in cond conversion with error " + caught.Message);
				this.m_intError=-1;
				p_OleDbCommand = null;
				p_ds = null;
				p_da = null;
				return;
			}
			if (this.m_intError==0)
			{
				for (intCurrRec = 0; intCurrRec <= p_ds.Tables["source_cond"].Rows.Count-1;intCurrRec++)
				{
					strbiosumcondid="";
					strbiosumid="";
					strplotid="";
					strcondid="";
					strcondnum="";
					strcondprop="";
					strlandclcd="";
					strfortypcd="";
					strglc="";
					strowncd="";
					strowngrpcd="";
					strreservcd="";
					strsiteclcd="";
					
					strsicond="";
					strsisp="";
					strslope="";
					straspect="";
					strstdage="";
					strstdszcd="";
					strhabtypcd1="";
					stradforcd="";
					strqmdtotcm="";
					stracres="";
					strunitcd="";
					strvollocgrp="";
					strtpa="";
					
					strvolacgrsstemttlft3="";
					strvolacgrsft3="";
					strvolacgrsstemsawlogft3="";
					
					strprenotcalc_yn="";
					strpretotflamesev="";
					strpretotflamemod="";
					strpretorchidx="";
					strprecrownidx="";
					strprecanopyht="";
					strprecanopydensity="";
					strpremortbasev="";
					strpremortbamod="";
					strpremortvolsev="";
					strpremortvolmod="";
					strqmdhwdcm = "";
					strqmdswdcm = "";
					strbaft2ac = "";
					strprefiretypesev="";
					strprefiretypemod="";
					
					strtopht="";
					strsdi="";
					strccf="";

					this.m_frmTherm.Increment(intCurrRec + 1);
					strcondid = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["idb_condition"]);
					strplotid = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["idb_plot"]); 
					strcondnum = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["cond_num"]); 
					strlandclcd = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["glc_group"]);
					if (strlandclcd.Trim().Length > 0)
					{
						switch (strlandclcd.Trim())
						{
							case "F":
								strlandclcd="1";
								break;
							case "NF":
								strlandclcd="2";
								break;
							case "CW":
								strlandclcd="4";
								break;
							default:
								strlandclcd="2";
								break;
						}

					}
                    strfortypcd = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["for_type"]); 
					strglc = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["glc"]);
					strowngrpcd = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["own_group"]);
					if (strowngrpcd.Trim().Length > 0)
					{
						switch (strowngrpcd.Trim())
						{
							case "1":
						        strowngrpcd = "10";
								break;
							case "2":
								strowngrpcd = "20";
								break;
							case "3":
								strowngrpcd = "30";
								break;
							case "4":
								strowngrpcd = "40";
								break;
							default:
								break;
						}
					}
					strsiteclcd = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["site_class_fia"]); 
                    strslope = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["slope"]);
                    strstdage = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["stand_age"]);
					strstdszcd = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["stand_size_class"]);
                    stradforcd = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["forest_or_blm_district"]);
					strqmdtotcm = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["qmd_tot"]);
					strqmdhwdcm = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["qmd_hwd"]);
					strqmdswdcm = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["qmd_swd"]);
					stracres = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["acres"]);
					strtpa = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["tpa"]);
					strbaft2ac = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["ba_ft2_ac"]);
                    strvolacgrsstemttlft3 = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["vol_ac_grs_stem_ttl_ft3"]);
					strvolacgrsft3 = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["vol_ac_grs_ft3"]);
					strvolacgrsstemsawlogft3 = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["vol_ac_grs_stem_sawlog_ft3"]);
                    strsdi = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["sdi"]);
					strccf = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["ccf"]);
					strtopht = Convert.ToString(p_ds.Tables["source_cond"].Rows[intCurrRec]["top_ht"]);

					strSQL = "select biosum_plot_id,half_state,unitcd " + 
						     "from target_plot where " + 
						     "idb_plot_id = " + strplotid + ";";

					p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
					p_OleDbCommand.CommandText = strSQL;
					p_da.SelectCommand = p_OleDbCommand;
					p_da.Fill(p_ds, "target_plot");
					if (p_ds.Tables["target_plot"].Rows.Count == 1)
					{
					   strbiosumid = Convert.ToString(p_ds.Tables["target_plot"].Rows[0]["biosum_plot_id"]);
                       strbiosumcondid = Convert.ToString(p_ds.Tables["target_plot"].Rows[0]["biosum_plot_id"]);
					   strbiosumcondid = strbiosumcondid.Trim();
					   strbiosumcondid += strcondnum.Trim();
					   strvollocgrp = Convert.ToString(p_ds.Tables["target_plot"].Rows[0]["half_state"]);
					   strvollocgrp = "S2SL" + strvollocgrp.Trim();
					   strunitcd = Convert.ToString(p_ds.Tables["target_plot"].Rows[0]["unitcd"]); 
					}
					else
					{
						MessageBox.Show("Fatal Error!!! Could not find plot record for cond id " + strcondid);
						this.m_txtStreamWriter.Write("!!!!Fatal Error!!! Could not find plot record for cond id " + strcondid + "\n");
						this.m_intError=-1;
						break;
					}
					p_ds.Tables["target_plot"].Clear();
					try
					{
						strSQL = "select pre_not_calc," + 
							"pre_sev_flam_lgth,pre_mod_flam_lgth,pre_sev_type,pre_mod_type," +
							"pre_torch_ind,pre_crown_ind,pre_cnpy_base_hgth,pre_cnpy_bulk_den," +
							"pre_sev_cf_mort,pre_mod_cf_mort,pre_torch_ind,pre_crown_ind," + 
							"pre_cnpy_base_hgth,pre_cnpy_bulk_den from source_ffe where " + 
							"idb_condition = " + strcondid + ";";

						p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
						p_OleDbCommand.CommandText = strSQL;
						p_da.SelectCommand = p_OleDbCommand;
						p_da.Fill(p_ds, "source_ffe");
					}
					catch (Exception caught)
					{
						MessageBox.Show(caught.Message);
					}
					if (p_ds.Tables["source_ffe"].Rows.Count > 0)
					{
						strprenotcalc_yn=Convert.ToString(p_ds.Tables["source_ffe"].Rows[0]["pre_not_calc"]);
						if (strprenotcalc_yn.IndexOf("F",0) >=0 )
						{
						   strprenotcalc_yn = "N";
						}
						else
						{
							strprenotcalc_yn = "Y";
						}
						strpretorchidx = Convert.ToString(p_ds.Tables["source_ffe"].Rows[0]["pre_torch_ind"]);
						strprecrownidx = Convert.ToString(p_ds.Tables["source_ffe"].Rows[0]["pre_crown_ind"]);
						strprecanopyht = Convert.ToString(p_ds.Tables["source_ffe"].Rows[0]["pre_cnpy_base_hgth"]);
						strprecanopydensity = Convert.ToString(p_ds.Tables["source_ffe"].Rows[0]["pre_cnpy_bulk_den"]);
						strpretotflamesev = Convert.ToString(p_ds.Tables["source_ffe"].Rows[0]["pre_sev_flam_lgth"]);
						strpretotflamemod = Convert.ToString(p_ds.Tables["source_ffe"].Rows[0]["pre_mod_flam_lgth"]);
						strprefiretypesev = Convert.ToString(p_ds.Tables["source_ffe"].Rows[0]["pre_sev_type"]);
						strprefiretypemod = Convert.ToString(p_ds.Tables["source_ffe"].Rows[0]["pre_mod_type"]);
						strpremortbasev = Convert.ToString(p_ds.Tables["source_ffe"].Rows[0]["pre_sev_cf_mort"]);
						strpremortbamod = Convert.ToString(p_ds.Tables["source_ffe"].Rows[0]["pre_mod_cf_mort"]);
					}
					else
					{
						this.m_txtStreamWriter.Write("Data Error: Could Not Find Cond Id = " + strcondid + " in ca/or ffe table!!! \n");
					}
					p_ds.Tables["source_ffe"].Clear();

                     

					strSQL = "select plant_assoc_code,cond_wt,owner,reserved_type,site_species," + 
						"aspect_deg,site_index_fia from pnw_idb_cond where " + 
						"cond_id = " + strcondid + ";";

					p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
					p_OleDbCommand.CommandText = strSQL;
					p_da.SelectCommand = p_OleDbCommand;
					p_da.Fill(p_ds, "pnw_idb_cond");
					if (p_ds.Tables["pnw_idb_cond"].Rows.Count == 1)
					{
						strhabtypcd1 = Convert.ToString(p_ds.Tables["pnw_idb_cond"].Rows[0]["plant_assoc_code"]);
						if (strhabtypcd1.Trim().Length == 0) strhabtypcd1 = "";
						strcondprop = Convert.ToString(p_ds.Tables["pnw_idb_cond"].Rows[0]["cond_wt"]); 
						if (strcondprop.Trim().Length == 0) strcondprop = "";
   					    strowncd = Convert.ToString(p_ds.Tables["pnw_idb_cond"].Rows[0]["owner"]); 
						if (strowncd.Trim().Length == 0) strowncd = "";
						strreservcd = Convert.ToString(p_ds.Tables["pnw_idb_cond"].Rows[0]["reserved_type"]); 
						if (strreservcd.Trim().Length > 0)
						{
							switch (strreservcd.Trim())
							{
								case "1":
									strreservcd = "1";
									break;
								case "5":
									strreservcd = "1";
									break;
								default:
									strreservcd = "0";
									break;
							}
						}
	                    strsisp = Convert.ToString(p_ds.Tables["pnw_idb_cond"].Rows[0]["site_species"]);
						if (strsisp.Trim().Length == 0) strsisp="";
                        straspect = Convert.ToString(p_ds.Tables["pnw_idb_cond"].Rows[0]["aspect_deg"]);
						if (straspect.Trim().Length == 0) straspect="";
						strsicond = Convert.ToString(p_ds.Tables["pnw_idb_cond"].Rows[0]["site_index_fia"]);
						if (strsicond.Trim().Length == 0) strsicond="";
					}
					else
					{
						strhabtypcd1="";
						strcondprop="";
						strowncd="";
						strreservcd="0";
						strsisp="";
						straspect="";
						this.m_txtStreamWriter.Write("Data Error: Could Not Find Cond Id = " + strcondid + " in pnw idb cond table!!! \n");
					}
					p_ds.Tables["pnw_idb_cond"].Clear();

					strDbFields = "INSERT INTO target_cond (biosum_cond_id,idb_cond_id,biosum_plot_id,idb_plot_id";
					strValues=" VALUES ('" + strbiosumcondid + "'," + strcondid + ",'" + strbiosumid + "'," + strplotid ;
					
					if (strcondnum.Trim().Length > 0)
					{
						strDbFields += ",condid";
						strValues += "," + strcondnum;
					}
					if (strcondprop.Trim().Length > 0)
					{
						strDbFields += ",condprop";
						strValues += "," + strcondprop;
					}
					if (strlandclcd.Trim().Length > 0)
					{
						strDbFields += ",landclcd";
						strValues += "," + strlandclcd;
					}
					if (strfortypcd.Trim().Length > 0)
					{
						strDbFields += ",fortypcd";
						strValues += "," + strfortypcd;
					}
                    
					if (strglc.Trim().Length > 0)
					{
						strDbFields += ",ground_land_class_pnw";
						strValues += ",'" + strglc + "'";
					}
					if (strowncd.Trim().Length > 0)
					{
						strDbFields += ",owncd";
						strValues += "," + strowncd;
					}
					if (strowngrpcd.Trim().Length > 0)
					{
						strDbFields += ",owngrpcd";
						strValues += "," + strowngrpcd;
					}
					if (strreservcd.Trim().Length > 0)
					{
						strDbFields += ",reservcd";
						strValues += "," + strreservcd;
					}
					if (strsiteclcd.Trim().Length > 0)
					{
						strDbFields += ",siteclcd";
						strValues += "," + strsiteclcd;
					}

					if (strsicond.Trim().Length > 0)
					{
						strDbFields += ",sicond";
						strValues += "," + strsicond;
					}
					if (strsisp.Trim().Length > 0)
					{
						strDbFields += ",sisp";
						strValues += "," + strsisp;
					}
					if (strslope.Trim().Length > 0)
					{
						strDbFields += ",slope";
						strValues += "," + strslope;
					}
					if (straspect.Trim().Length > 0)
					{
						strDbFields += ",aspect";
						strValues += "," + straspect;
					}
					if (strstdage.Trim().Length > 0)
					{
						strDbFields += ",stdage";
						strValues += "," + strstdage;
					}
					if (strstdszcd.Trim().Length > 0)
					{
						strDbFields += ",stdszcd";
						strValues += "," + strstdszcd;
					}
					if (strhabtypcd1.Trim().Length > 0)
					{
						strDbFields += ",habtypcd1";
						strValues += ",'" + strhabtypcd1 + "'";
					}
					if (stradforcd.Trim().Length > 0)
					{
						strDbFields += ",adforcd";
						strValues += "," + stradforcd;
					}
					if (strqmdtotcm.Trim().Length > 0)
					{
						strDbFields += ",qmd_tot_cm";
						strValues += "," + strqmdtotcm;
					}
					if (strqmdhwdcm.Trim().Length > 0)
					{
						strDbFields += ",qmd_hwd_cm";
						strValues += "," + strqmdhwdcm;
					}
					if (strqmdswdcm.Trim().Length > 0)
					{
						strDbFields += ",qmd_swd_cm";
						strValues += "," + strqmdswdcm;
					}
					if (stracres.Trim().Length > 0)
					{
						strDbFields += ",acres";
						strValues += "," + stracres;
					}
					if (strunitcd.Trim().Length > 0)
					{
						strDbFields += ",unitcd";
						strValues += "," + strunitcd;
					}
					if (strvollocgrp.Trim().Length > 0)
					{
						strDbFields += ",vol_loc_grp";
						strValues += ",'" + strvollocgrp + "'";
					}
					if (strtpa.Trim().Length > 0)
					{
						strDbFields += ",tpacurr";
						strValues += "," + strtpa;
					}
					if (strbaft2ac.Trim().Length > 0)
					{
						strDbFields += ",ba_ft2_ac";
						strValues += "," + strbaft2ac;
					}
					if (strvolacgrsstemttlft3.Trim().Length > 0)
					{
						strDbFields += ",vol_ac_grs_stem_ttl_ft3";
						strValues += "," + strvolacgrsstemttlft3;
					}
					if (strvolacgrsft3.Trim().Length > 0)
					{
						strDbFields += ",vol_ac_grs_ft3";
						strValues += "," + strvolacgrsft3;
					}
					if (strvolacgrsstemsawlogft3.Trim().Length > 0)
					{
						strDbFields += ",volcsgrs";
						strValues += "," + strvolacgrsstemsawlogft3;
					}
					if (strprenotcalc_yn.Trim().Length > 0)
					{
						strDbFields += ",pre_not_calc_yn";
						strValues += ",'" + strprenotcalc_yn + "'";
					}
					else
					{
						strDbFields += ",pre_not_calc_yn";
						strValues += ",'Y'";

					}
					if (strpretotflamemod.Trim().Length > 0)
					{
						strDbFields += ",pre_tot_flame_mod";
						strValues += "," + strpretotflamemod;
					}
					if (strprefiretypesev.Trim().Length > 0)
					{
						strDbFields += ",pre_fire_type_sev";
						strValues += ",'" + strprefiretypesev + "'";
					}
					if (strprefiretypemod.Trim().Length > 0)
					{
						strDbFields += ",pre_fire_type_mod";
						strValues += ",'" + strprefiretypemod + "'";
					}
					if (strpretorchidx.Trim().Length > 0)
					{
						strDbFields += ",pre_torch_index";
						strValues += "," + strpretorchidx;
					}
					if (strprecrownidx.Trim().Length > 0)
					{
						strDbFields += ",pre_crown_index";
						strValues += "," + strprecrownidx;
					}
					if (strprecanopyht.Trim().Length > 0)
					{
						strDbFields += ",pre_canopy_ht";
						strValues += "," + strprecanopyht;
					}
					if (strprecanopydensity.Trim().Length > 0)
					{
						strDbFields += ",pre_canopy_density";
						strValues += "," + strprecanopydensity;
					}
					if (strpremortbasev.Trim().Length > 0)
					{
						strDbFields += ",pre_mortality_ba_sev";
						strValues += "," + strpremortbasev;
					}
					if (strpremortbamod.Trim().Length > 0)
					{
						strDbFields += ",pre_mortality_ba_mod";
						strValues += "," + strpremortbamod;
					}
					if (strpremortvolsev.Trim().Length > 0)
					{
						strDbFields += ",pre_mortality_vol_sev";
						strValues += "," + strpremortvolsev;
					}
					if (strpremortvolmod.Trim().Length > 0)
					{
						strDbFields += ",pre_mortality_vol_mod";
						strValues += "," + strpremortvolmod;
					}
					
					if (strsdi.Trim().Length > 0)
					{
						strDbFields += ",sdi";
						strValues += "," + strsdi;
					}
					if (strccf.Trim().Length > 0)
					{
						strDbFields += ",ccf";
						strValues += "," + strccf;
					}
					if (strtopht.Trim().Length > 0)
					{
						strDbFields += ",topht";
						strValues += "," + strtopht;
					}


					strDbFields += ")";
					strValues += ");";
					strSQL = strDbFields + strValues;

					this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strSQL);
					if (this.m_ado.m_intError < 0)
					{
						this.m_txtStreamWriter.Write("Data Error: insert sql error for sql statement " + strSQL + "\n");
					}
					System.Windows.Forms.Application.DoEvents();
					if (this.m_frmTherm.AbortProcess == true) break;
				}
			}
			p_OleDbCommand = null;
			p_ds.Clear();
			p_da.Dispose();
			p_ds = null;
			p_da = null;
			this.m_txtStreamWriter.Write("Cond Records Processed: " + intCurrRec.ToString() + "\n");

		}
		private void convert_ffe()
		{
			System.Data.OleDb.OleDbDataAdapter p_da;
			System.Data.DataSet p_ds;
			System.Data.OleDb.OleDbCommand p_OleDbCommand;
			p_da = new System.Data.OleDb.OleDbDataAdapter();
			p_ds = new DataSet("tempds");
			this.m_intError=0;
			this.m_strError = "";
			string strSQL="";
			int intCurrRec=0;

			string strbiosumcondid,strcondid;
			string strrx;
			string strpostnotcalc_yn,strposttotflamesev;
			string strposttotflamemod, strposttorchidx, strpostcrownidx;
			string strpostcanopyht, strpostcanopydensity, strpostmortbasev;
			string strpostmortbamod, strpostmortvolsev, strpostmortvolmod;
			string strpostfiretypesev,strpostfiretypemod;

			string strDbFields="";
			string strValues="";
			
			this.m_intError=0;
			this.m_txtStreamWriter.Write("\n\nConverting FFE Data\n");
			this.m_txtStreamWriter.Write("--------------------\n");

			try
			{
				strSQL = "DELETE * FROM TARGET_FFE";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, strSQL);
				strSQL = "SELECT * FROM SOURCE_FFE";
				m_frmTherm.AbortProcess = false;
				m_frmTherm.lblMsg.Text = "Converting FFE Data";
				m_frmTherm.lblMsg.Visible = true;
				m_frmTherm.lblMsg.Refresh();
				this.m_frmTherm.progressBar1.Maximum = Convert.ToInt32(this.m_ado.getRecordCount("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.strTempMDBFile + ";User Id=admin;Password=;","select count(*) from source_ffe","source_ffe"));
				this.m_txtStreamWriter.Write("source ffe record count=" + Convert.ToString(this.m_frmTherm.progressBar1.Maximum) + "\n");
				p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
				p_OleDbCommand.CommandText = strSQL;
				p_da.SelectCommand = p_OleDbCommand;
				p_da.Fill(p_ds, "source_ffe");
			}
			catch (Exception caught)
			{
				MessageBox.Show("Abort Process: Error in cond conversion with error " + caught.Message);
				this.m_intError=-1;
				p_OleDbCommand = null;
				p_ds = null;
				p_da = null;
				return;
			}
			if (this.m_intError==0)
			{
				for (intCurrRec = 0; intCurrRec <= p_ds.Tables["source_ffe"].Rows.Count-1;intCurrRec++)
				{
					strbiosumcondid="";
					strcondid="";
					strrx="";
					strpostnotcalc_yn="";
					strposttotflamesev="";
					strposttotflamemod="";
					strposttorchidx="";
					strpostcrownidx="";
					strpostcanopyht="";
					strpostcanopydensity="";
					strpostmortbasev="";
					strpostmortbamod="";
					strpostmortvolsev="";
					strpostmortvolmod="";
					strpostfiretypesev="";
					strpostfiretypemod="";

					this.m_frmTherm.Increment(intCurrRec + 1);
					strcondid = Convert.ToString(p_ds.Tables["source_ffe"].Rows[intCurrRec]["idb_condition"]);
					strrx = Convert.ToString(p_ds.Tables["source_ffe"].Rows[intCurrRec]["rx"]); 
					strposttotflamesev = Convert.ToString(p_ds.Tables["source_ffe"].Rows[intCurrRec]["sev_flam_lgth"]); 
					strposttotflamemod = Convert.ToString(p_ds.Tables["source_ffe"].Rows[intCurrRec]["mod_flam_lgth"]); 
     				strpostfiretypesev = Convert.ToString(p_ds.Tables["source_ffe"].Rows[intCurrRec]["sev_type"]); 
     				strpostfiretypemod = Convert.ToString(p_ds.Tables["source_ffe"].Rows[intCurrRec]["mod_type"]);
     				strposttorchidx = Convert.ToString(p_ds.Tables["source_ffe"].Rows[intCurrRec]["torch_ind"]); 
					strpostcrownidx = Convert.ToString(p_ds.Tables["source_ffe"].Rows[intCurrRec]["crown_ind"]);
     				strpostcanopyht = Convert.ToString(p_ds.Tables["source_ffe"].Rows[intCurrRec]["cnpy_base_hgth"]); 
                    strpostcanopydensity = Convert.ToString(p_ds.Tables["source_ffe"].Rows[intCurrRec]["cnpy_bulk_den"]); 
					strpostmortvolsev = Convert.ToString(p_ds.Tables["source_ffe"].Rows[intCurrRec]["sev_cf_mort"]); 
					strpostmortvolmod = Convert.ToString(p_ds.Tables["source_ffe"].Rows[intCurrRec]["mod_cf_mort"]); 
					strpostnotcalc_yn = Convert.ToString(p_ds.Tables["source_ffe"].Rows[intCurrRec]["post_not_calc"]); 
					if (strpostnotcalc_yn.IndexOf("F",0) >=0 )
					{
						strpostnotcalc_yn = "N";
					}
					else
					{
						strpostnotcalc_yn = "Y";
					}

					strSQL = "select biosum_cond_id from target_cond where " + 
						"idb_cond_id = " + strcondid + ";";

					p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
					p_OleDbCommand.CommandText = strSQL;
					p_da.SelectCommand = p_OleDbCommand;
					p_da.Fill(p_ds, "target_cond");
					if (p_ds.Tables["target_cond"].Rows.Count == 1)
					{
						strbiosumcondid = Convert.ToString(p_ds.Tables["target_cond"].Rows[0]["biosum_cond_id"]);
					}
					else
					{
						MessageBox.Show("Fatal Error!!! Could not find cond record for cond id " + strcondid);
						this.m_txtStreamWriter.Write("!!!!Fatal Error!!! Could not find cond record for cond id " + strcondid + "\n");
						this.m_intError=-1;
						break;
					}
					p_ds.Tables["target_cond"].Clear();

					strDbFields = "INSERT INTO target_ffe (biosum_cond_id";
					strValues=" VALUES ('" + strbiosumcondid + "'";
					
					if (strrx.Trim().Length > 0)
					{
						strDbFields += ",rx";
						strValues += ",'" + strrx + "'";
					}
					if (strpostnotcalc_yn.Trim().Length > 0)
					{
						strDbFields += ",post_not_calc_yn";
						strValues += ",'" + strpostnotcalc_yn + "'";
					}
					if (strposttotflamesev.Trim().Length > 0)
					{
						strDbFields += ",post_tot_flame_sev";
						strValues += "," + strposttotflamesev;
					}
					if (strposttotflamemod.Trim().Length > 0)
					{
						strDbFields += ",post_tot_flame_mod";
						strValues += "," + strposttotflamemod;
					}
                    
					if (strpostfiretypesev.Trim().Length > 0)
					{
						strDbFields += ",post_fire_type_sev";
						strValues += ",'" + strpostfiretypesev + "'";
					}
					if (strpostfiretypemod.Trim().Length > 0)
					{
						strDbFields += ",post_fire_type_mod";
						strValues += ",'" + strpostfiretypemod + "'";
					}
					if (strposttorchidx.Trim().Length > 0)
					{
						strDbFields += ",post_torch_index";
						strValues += "," + strposttorchidx;
					}
					if (strpostcrownidx.Trim().Length > 0)
					{
						strDbFields += ",post_crown_index";
						strValues += "," + strpostcrownidx;
					}
					if (strpostcanopyht.Trim().Length > 0)
					{
						strDbFields += ",post_canopy_ht";
						strValues += "," + strpostcanopyht;
					}
					if (strpostcanopydensity.Trim().Length > 0)
					{
						strDbFields += ",post_canopy_density";
						strValues += "," + strpostcanopydensity;
					}
					if (strpostmortvolsev.Trim().Length > 0)
					{
						strDbFields += ",post_mortality_vol_sev";
						strValues += "," + strpostmortvolsev;
					}
					if (strpostmortvolmod.Trim().Length > 0)
					{
						strDbFields += ",post_mortality_vol_mod";
						strValues += "," + strpostmortvolmod;
					}
					strDbFields += ")";
					strValues += ");";
					strSQL = strDbFields + strValues;

					this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strSQL);
					if (this.m_ado.m_intError < 0)
					{
						this.m_txtStreamWriter.Write("Data Error: insert sql error for sql statement " + strSQL + "\n");
					}
					System.Windows.Forms.Application.DoEvents();
					if (this.m_frmTherm.AbortProcess == true) break;
				}
			}
			p_OleDbCommand = null;
			p_ds.Clear();
			p_da.Dispose();
			p_ds = null;
			p_da = null;
			this.m_txtStreamWriter.Write("ffe Records Processed: " + intCurrRec.ToString() + "\n");

		}
		private void convert_treediamgroups()
		{
			System.Data.OleDb.OleDbDataAdapter p_da;
			System.Data.DataSet p_ds;
			System.Data.OleDb.OleDbCommand p_OleDbCommand;
			p_da = new System.Data.OleDb.OleDbDataAdapter();
			p_ds = new DataSet("tempds");
			this.m_intError=0;
			this.m_strError = "";
			string strSQL="";
			int intCurrRec=0;

			string strdiamid,strdiamclass;

			string strDbFields="";
			string strValues="";
			
			this.m_intError=0;
			this.m_txtStreamWriter.Write("\n\nConverting Tree Diameter Groups Data\n");
			this.m_txtStreamWriter.Write("--------------------\n");

			try
			{
				strSQL = "DELETE * FROM TARGET_TREE_DIAM_GROUPS";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, strSQL);
				strSQL = "SELECT * FROM SOURCE_TREE_DIAM_GROUPS";
				m_frmTherm.AbortProcess = false;
				m_frmTherm.lblMsg.Text = "Converting Tree Diameter Groups Data";
				m_frmTherm.lblMsg.Visible = true;
				m_frmTherm.lblMsg.Refresh();
				this.m_frmTherm.progressBar1.Maximum = Convert.ToInt32(this.m_ado.getRecordCount("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.strTempMDBFile + ";User Id=admin;Password=;","select count(*) from source_tree_diam_groups","source_tree_diam_groups"));
				this.m_txtStreamWriter.Write("source ref_diameter_classes record count=" + Convert.ToString(this.m_frmTherm.progressBar1.Maximum) + "\n");
				p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
				p_OleDbCommand.CommandText = strSQL;
				p_da.SelectCommand = p_OleDbCommand;
				p_da.Fill(p_ds, "source_tree_diam_groups");
			}
			catch (Exception caught)
			{
				MessageBox.Show("Abort Process: Error in cond conversion with error " + caught.Message);
				this.m_intError=-1;
				p_OleDbCommand = null;
				p_ds = null;
				p_da = null;
				return;
			}
			if (this.m_intError==0)
			{
				for (intCurrRec = 0; intCurrRec <= p_ds.Tables["source_tree_diam_groups"].Rows.Count-1;intCurrRec++)
				{
					strdiamid = "";
					strdiamclass="";
					this.m_frmTherm.Increment(intCurrRec + 1);
					strdiamid = Convert.ToString(p_ds.Tables["source_tree_diam_groups"].Rows[intCurrRec]["diameterid"]);
					strdiamclass = Convert.ToString(p_ds.Tables["source_tree_diam_groups"].Rows[intCurrRec]["diam_class"]); 

					strDbFields = "INSERT INTO target_tree_diam_groups (";
					strValues=" VALUES (";
					
					if (strdiamid.Trim().Length > 0)
					{
						strDbFields += "diam_group";
						strValues += strdiamid;
					}
					if (strdiamclass.Trim().Length > 0)
					{
						strDbFields += ",diam_class";
						strValues += ",'" + strdiamclass + "'";
					}
					strDbFields += ")";
					strValues += ");";
					strSQL = strDbFields + strValues;

					this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strSQL);
					if (this.m_ado.m_intError < 0)
					{
						this.m_txtStreamWriter.Write("Data Error: insert sql error for sql statement " + strSQL + "\n");
					}
					System.Windows.Forms.Application.DoEvents();
					if (this.m_frmTherm.AbortProcess == true) break;
				}
			}
			p_OleDbCommand = null;
			p_ds.Clear();
			p_da.Dispose();
			p_ds = null;
			p_da = null;
			this.m_txtStreamWriter.Write("tree_diam_groups Records Processed: " + intCurrRec.ToString() + "\n");

		}
		private void convert_rx()
		{
			System.Data.OleDb.OleDbDataAdapter p_da;
			System.Data.DataSet p_ds;
			System.Data.OleDb.OleDbCommand p_OleDbCommand;
			p_da = new System.Data.OleDb.OleDbDataAdapter();
			p_ds = new DataSet("tempds");
			
			this.m_intError=0;
			this.m_strError = "";
			string strSQL="";
			int intCurrRec=0;

			string strrx,strrxdesc,strrxintensity;

			string strDbFields="";
			string strValues="";
			
			this.m_intError=0;
			this.m_txtStreamWriter.Write("\n\nTreatment Prescriptions (Rx) Data\n");
			this.m_txtStreamWriter.Write("-----------------------------------------\n");

			try
			{
				strSQL = "DELETE * FROM TARGET_RX";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, strSQL);
				strSQL = "DELETE * FROM TARGET_RX_INTENSITY";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, strSQL);
				strSQL = "SELECT * FROM SOURCE_RX";
				m_frmTherm.AbortProcess = false;
				m_frmTherm.lblMsg.Text = "Converting Treatment Prescriptions (RX) Data";
				m_frmTherm.lblMsg.Visible = true;
				m_frmTherm.lblMsg.Refresh();
				this.m_frmTherm.progressBar1.Maximum = Convert.ToInt32(this.m_ado.getRecordCount("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.strTempMDBFile + ";User Id=admin;Password=;","select count(*) from source_rx","source_rx"));
				this.m_txtStreamWriter.Write("source rx record count=" + Convert.ToString(this.m_frmTherm.progressBar1.Maximum) + "\n");
				p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
				p_OleDbCommand.CommandText = strSQL;
				p_da.SelectCommand = p_OleDbCommand;
				p_da.Fill(p_ds, "source_rx");
			}
			catch (Exception caught)
			{
				MessageBox.Show("Abort Process: Error in rx conversion with error " + caught.Message);
				this.m_intError=-1;
				p_OleDbCommand = null;
				p_ds = null;
				p_da = null;
				return;
			}
			if (this.m_intError==0)
			{
				for (intCurrRec = 0; intCurrRec <= p_ds.Tables["source_rx"].Rows.Count-1;intCurrRec++)
				{
					strrx = "";
					strrxintensity="";
					strrxdesc = "";
					this.m_frmTherm.Increment(intCurrRec + 1);
					strrx = Convert.ToString(p_ds.Tables["source_rx"].Rows[intCurrRec]["rx"]);
					strrxintensity = Convert.ToString(p_ds.Tables["source_rx"].Rows[intCurrRec]["intensity"]); 
					strrxdesc = Convert.ToString(p_ds.Tables["source_rx"].Rows[intCurrRec]["description"]); 
					strrxdesc = this.m_ado.FixString(strrxdesc,"'","''");
					strDbFields = "INSERT INTO target_rx (";
					strValues=" VALUES (";
					
					if (strrx.Trim().Length > 0)
					{
						strDbFields += "rx";
						strValues += "'" + strrx + "'";
					}
					if (strrxdesc.Trim().Length > 0)
					{
						strDbFields += ",description";
						strValues += ",'" + strrxdesc + "'";
					}
					strDbFields += ")";
					strValues += ");";
					strSQL = strDbFields + strValues;

					this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strSQL);
					if (this.m_ado.m_intError < 0)
					{
						this.m_txtStreamWriter.Write("Data Error: insert sql error for sql statement " + strSQL + "\n");
					}

					strDbFields = "INSERT INTO target_rx_intensity (scenario_id";
					strValues=" VALUES ('scenario1'";
					
					if (strrx.Trim().Length > 0)
					{
						strDbFields += ",rx";
						strValues += ",'" + strrx + "'";
					}
					if (strrxintensity.Trim().Length > 0)
					{
						strDbFields += ",rx_intensity";
						strValues += "," + strrxintensity;
					}
					strDbFields += ")";
					strValues += ");";
					strSQL = strDbFields + strValues;

					this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strSQL);
					if (this.m_ado.m_intError < 0)
					{
						this.m_txtStreamWriter.Write("Data Error: insert sql error for sql statement " + strSQL + "\n");
					}
					System.Windows.Forms.Application.DoEvents();
					if (this.m_frmTherm.AbortProcess == true) break;
				}
			}
			p_OleDbCommand = null;
			p_ds.Clear();
			p_da.Dispose();
			p_ds = null;
			p_da = null;
			this.m_txtStreamWriter.Write("rx Records Processed: " + intCurrRec.ToString() + "\n");

		}
		private void convert_treespeciesgroups()
		{
			System.Data.OleDb.OleDbDataAdapter p_da;
			System.Data.DataSet p_ds;
			System.Data.OleDb.OleDbCommand p_OleDbCommand;
			p_da = new System.Data.OleDb.OleDbDataAdapter();
			p_ds = new DataSet("tempds");
			this.m_intError=0;
			this.m_strError = "";
			string strSQL="";
			int intCurrRec=0;

			string strspeciesgroup,strspecieslist,strspecieslabel;

			string strDbFields="";
			string strValues="";
			
			this.m_intError=0;
			this.m_txtStreamWriter.Write("\n\nTree Species Groups Data\n");
			this.m_txtStreamWriter.Write("--------------------\n");

			try
			{
				strSQL = "DELETE * FROM TARGET_TREE_SPECIES_GROUPS";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, strSQL);
				strSQL = "SELECT * FROM SOURCE_TREE_SPECIES_GROUPS";
				m_frmTherm.AbortProcess = false;
				m_frmTherm.lblMsg.Text = "Converting Tree Species Groups Data";
				m_frmTherm.lblMsg.Visible = true;
				m_frmTherm.lblMsg.Refresh();
				this.m_frmTherm.progressBar1.Maximum = Convert.ToInt32(this.m_ado.getRecordCount("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.strTempMDBFile + ";User Id=admin;Password=;","select count(*) from source_tree_species_groups","source_tree_species_groups"));
				this.m_txtStreamWriter.Write("source ref_spp_grp record count=" + Convert.ToString(this.m_frmTherm.progressBar1.Maximum) + "\n");
				p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
				p_OleDbCommand.CommandText = strSQL;
				p_da.SelectCommand = p_OleDbCommand;
				p_da.Fill(p_ds, "source_tree_species_groups");
			}
			catch (Exception caught)
			{
				MessageBox.Show("Abort Process: Error in tree species groups conversion with error " + caught.Message);
				this.m_intError=-1;
				p_OleDbCommand = null;
				p_ds = null;
				p_da = null;
				return;
			}
			if (this.m_intError==0)
			{
				for (intCurrRec = 0; intCurrRec <= p_ds.Tables["source_tree_species_groups"].Rows.Count-1;intCurrRec++)
				{
					strspeciesgroup = "";
					strspecieslist="";
					strspecieslabel="";
					this.m_frmTherm.Increment(intCurrRec + 1);
					strspeciesgroup = Convert.ToString(p_ds.Tables["source_tree_species_groups"].Rows[intCurrRec]["species_group"]);
					strspecieslist = Convert.ToString(p_ds.Tables["source_tree_species_groups"].Rows[intCurrRec]["species_list"]); 
					strspecieslabel = Convert.ToString(p_ds.Tables["source_tree_species_groups"].Rows[intCurrRec]["species_label"]); 

					strDbFields = "INSERT INTO target_tree_species_groups (";
					strValues=" VALUES (";
					
					if (strspeciesgroup.Trim().Length > 0)
					{
						strDbFields += "species_group";
						strValues += strspeciesgroup;
					}
					if (strspecieslist.Trim().Length > 0)
					{
						strDbFields += ",species_list";
						strValues += ",'" + strspecieslist + "'";
					}
					if (strspecieslabel.Trim().Length > 0)
					{
						strDbFields += ",species_label";
						strValues += ",'" + strspecieslabel + "'";
					}

					strDbFields += ")";
					strValues += ");";
					strSQL = strDbFields + strValues;

					this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strSQL);
					if (this.m_ado.m_intError < 0)
					{
						this.m_txtStreamWriter.Write("Data Error: insert sql error for sql statement " + strSQL + "\n");
					}
					System.Windows.Forms.Application.DoEvents();
					if (this.m_frmTherm.AbortProcess == true) break;
				}
			}
			p_OleDbCommand = null;
			p_ds.Clear();
			p_da.Dispose();
			p_ds = null;
			p_da = null;
			this.m_txtStreamWriter.Write("ref_spp_grp Records Processed: " + intCurrRec.ToString() + "\n");

		}
		private void convert_harvestcosts()
		{
				System.Data.OleDb.OleDbDataAdapter p_da;
				System.Data.DataSet p_ds;
				System.Data.OleDb.OleDbCommand p_OleDbCommand;
				p_da = new System.Data.OleDb.OleDbDataAdapter();
				p_ds = new DataSet("tempds");
				this.m_intError=0;
				this.m_strError = "";
				string strSQL="";
				int intCurrRec=0;

				string strbiosumcondid,strrx,strharvestcpa;
                string strwaterbarringroadscpa,strbrushcuttingcpa;
			    string strcompletecpa,strcondid;
			
				string strDbFields="";
				string strValues="";
			
				this.m_intError=0;
				this.m_txtStreamWriter.Write("\n\nHarvest Costs Data\n");
				this.m_txtStreamWriter.Write("-------------------------\n");

				try
				{
					strSQL = "DELETE * FROM TARGET_HARVEST_COSTS";
					this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, strSQL);
					strSQL = "SELECT * FROM SOURCE_HARVEST_COSTS";
					m_frmTherm.AbortProcess = false;
					m_frmTherm.lblMsg.Text = "Converting Harvest Costs Data";
					m_frmTherm.lblMsg.Visible = true;
					m_frmTherm.lblMsg.Refresh();
					this.m_frmTherm.progressBar1.Maximum = Convert.ToInt32(this.m_ado.getRecordCount("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.strTempMDBFile + ";User Id=admin;Password=;","select count(*) from source_harvest_costs","source_harvest_costs"));
					this.m_txtStreamWriter.Write("source harvest costs record count=" + Convert.ToString(this.m_frmTherm.progressBar1.Maximum) + "\n");
					p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
					p_OleDbCommand.CommandText = strSQL;
					p_da.SelectCommand = p_OleDbCommand;
					p_da.Fill(p_ds, "source_harvest_costs");
				}
				catch (Exception caught)
				{
					MessageBox.Show("Abort Process: Error in harvest costs conversion with error " + caught.Message);
					this.m_intError=-1;
					p_OleDbCommand = null;
					p_ds = null;
					p_da = null;
					return;
				}
				if (this.m_intError==0)
				{
					for (intCurrRec = 0; intCurrRec <= p_ds.Tables["source_harvest_costs"].Rows.Count-1;intCurrRec++)
					{
						strcondid="";
						strbiosumcondid = "";
						strharvestcpa="";
						strwaterbarringroadscpa="";
						strbrushcuttingcpa="";
						strcompletecpa="";
						strrx = "";
						this.m_frmTherm.Increment(intCurrRec + 1);
                        strcondid = Convert.ToString(p_ds.Tables["source_harvest_costs"].Rows[intCurrRec]["idb_condition"]);
						strrx = Convert.ToString(p_ds.Tables["source_harvest_costs"].Rows[intCurrRec]["rx"]);
						strharvestcpa = Convert.ToString(p_ds.Tables["source_harvest_costs"].Rows[intCurrRec]["total harvest cost ($/acre)"]); 
						strwaterbarringroadscpa = Convert.ToString(p_ds.Tables["source_harvest_costs"].Rows[intCurrRec]["water barring roads ($/acre)"]); 
                        strbrushcuttingcpa = Convert.ToString(p_ds.Tables["source_harvest_costs"].Rows[intCurrRec]["brush cutting ($/acre)"]); 
                        strcompletecpa = Convert.ToString(p_ds.Tables["source_harvest_costs"].Rows[intCurrRec]["completecstperac"]); 

						strSQL = "select biosum_cond_id from target_cond where " + 
							"idb_cond_id = " + strcondid + ";";

						p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
						p_OleDbCommand.CommandText = strSQL;
						p_da.SelectCommand = p_OleDbCommand;
						p_da.Fill(p_ds, "target_cond");
						if (p_ds.Tables["target_cond"].Rows.Count == 1)
						{
							strbiosumcondid = Convert.ToString(p_ds.Tables["target_cond"].Rows[0]["biosum_cond_id"]);
						}
						else
						{
							MessageBox.Show("Fatal Error!!! Could not find cond record for cond id " + strcondid);
							this.m_txtStreamWriter.Write("!!!!Fatal Error!!! Could not find cond record for cond id " + strcondid + "\n");
							this.m_intError=-1;
							break;
						}
						p_ds.Tables["target_cond"].Clear();

						strDbFields = "INSERT INTO target_harvest_costs (biosum_cond_id";
						strValues=" VALUES ('" + strbiosumcondid + "'";

						if (strrx.Trim().Length > 0)
						{
							strDbFields += ",rx";
							strValues += ",'" + strrx + "'";
						}

					
						if (strharvestcpa.Trim().Length > 0)
						{
							strDbFields += ",harvest_cpa";
							strValues += "," + strharvestcpa;
						}
						if (strwaterbarringroadscpa.Trim().Length > 0)
						{
							strDbFields += ",water_barring_roads_cpa";
							strValues += "," + strwaterbarringroadscpa;
						}

						if (strbrushcuttingcpa.Trim().Length > 0)
						{
							strDbFields += ",brush_cutting_cpa";
							strValues += "," + strbrushcuttingcpa;
						}

						if (strcompletecpa.Trim().Length > 0)
						{
							strDbFields += ",complete_cpa";
							strValues += "," + strcompletecpa;
						}

						strDbFields += ")";
						strValues += ");";
						strSQL = strDbFields + strValues;

						this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strSQL);
						if (this.m_ado.m_intError < 0)
						{
							this.m_txtStreamWriter.Write("Data Error: insert sql error for sql statement " + strSQL + "\n");
						}
						System.Windows.Forms.Application.DoEvents();
						if (this.m_frmTherm.AbortProcess == true) break;
					}
				}
				p_OleDbCommand = null;
				p_ds.Clear();
				p_da.Dispose();
				p_ds = null;
				p_da = null;
				this.m_txtStreamWriter.Write("harvest costs Records Processed: " + intCurrRec.ToString() + "\n");

		}
		private void convert_treevolval()
		{
				System.Data.OleDb.OleDbDataAdapter p_da;
				System.Data.DataSet p_ds;
				System.Data.OleDb.OleDbCommand p_OleDbCommand;
				p_da = new System.Data.OleDb.OleDbDataAdapter();
				p_ds = new DataSet("tempds");
				this.m_intError=0;
				this.m_strError = "";
				string strSQL="";
				int intCurrRec=0;

				string strbiosumcondid,strrx,strspeciesgroup;
        string strdiamid,strbiomassvolcf,strbiomasswtgt;
        string strmerchvolcf,strmerchwtgt, strmerchvaldpa;
        string strcondid;
			
				string strDbFields="";
				string strValues="";
			
				this.m_intError=0;
				this.m_txtStreamWriter.Write("\n\nTree Volume And Value By Species Diameter Groups\n");
				this.m_txtStreamWriter.Write("---------------------------------------------------------\n");

				try
				{
					strSQL = "DELETE * FROM TARGET_TREE_VOL_VAL_BY_SPECIES_DIAM_GROUPS";
					this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, strSQL);
					strSQL = "SELECT * FROM SOURCE_TREE_VOL_VAL_BY_SPECIES_DIAM_GROUPS";
					m_frmTherm.AbortProcess = false;
					m_frmTherm.lblMsg.Text = "Converting Tree Vol Val By Species Diam Groups Data";
					m_frmTherm.lblMsg.Visible = true;
					m_frmTherm.lblMsg.Refresh();
					this.m_frmTherm.progressBar1.Maximum = Convert.ToInt32(this.m_ado.getRecordCount("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.strTempMDBFile + ";User Id=admin;Password=;","select count(*) from source_tree_vol_val_by_species_diam_groups","source_tree_vol_val_by_species_diam_groups"));
					this.m_txtStreamWriter.Write("source tree vol val record count=" + Convert.ToString(this.m_frmTherm.progressBar1.Maximum) + "\n");
					p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
					p_OleDbCommand.CommandText = strSQL;
					p_da.SelectCommand = p_OleDbCommand;
					p_da.Fill(p_ds, "source_tree_vol_val_by_species_diam_groups");
				}
				catch (Exception caught)
				{
					MessageBox.Show("Abort Process: Error in tree vol val conversion with error " + caught.Message);
					this.m_intError=-1;
					p_OleDbCommand = null;
					p_ds = null;
					p_da = null;
					return;
				}
				if (this.m_intError==0)
				{
					for (intCurrRec = 0; intCurrRec <= p_ds.Tables["source_tree_vol_val_by_species_diam_groups"].Rows.Count-1;intCurrRec++)
					{
						strbiosumcondid="";
						strrx="";
						strspeciesgroup="";
                        strdiamid="";
                        strbiomassvolcf="";
                        strbiomasswtgt="";
                        strmerchvolcf="";
                        strmerchwtgt="";
                        strmerchvaldpa="";
                        strcondid="";

						this.m_frmTherm.Increment(intCurrRec + 1);
                        strcondid = Convert.ToString(p_ds.Tables["source_tree_vol_val_by_species_diam_groups"].Rows[intCurrRec]["idb_condition"]);
						strrx = Convert.ToString(p_ds.Tables["source_tree_vol_val_by_species_diam_groups"].Rows[intCurrRec]["rx"]);
						strspeciesgroup = Convert.ToString(p_ds.Tables["source_tree_vol_val_by_species_diam_groups"].Rows[intCurrRec]["species_group"]);
						strdiamid = Convert.ToString(p_ds.Tables["source_tree_vol_val_by_species_diam_groups"].Rows[intCurrRec]["diameterid"]);
						strbiomassvolcf = Convert.ToString(p_ds.Tables["source_tree_vol_val_by_species_diam_groups"].Rows[intCurrRec]["biomass volume(cf)"]);
						strbiomasswtgt = Convert.ToString(p_ds.Tables["source_tree_vol_val_by_species_diam_groups"].Rows[intCurrRec]["biomass weight(green tons)"]);
						strmerchvolcf = Convert.ToString(p_ds.Tables["source_tree_vol_val_by_species_diam_groups"].Rows[intCurrRec]["merchantable volume(cf)"]);
						strmerchwtgt  = Convert.ToString(p_ds.Tables["source_tree_vol_val_by_species_diam_groups"].Rows[intCurrRec]["merchantable weight(green tons)"]);
						strmerchvaldpa = Convert.ToString(p_ds.Tables["source_tree_vol_val_by_species_diam_groups"].Rows[intCurrRec]["merchantable value ($/ac)"]);
						
						strSQL = "select biosum_cond_id from target_cond where " + 
							"idb_cond_id = " + strcondid + ";";

						p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
						p_OleDbCommand.CommandText = strSQL;
						p_da.SelectCommand = p_OleDbCommand;
						p_da.Fill(p_ds, "target_cond");
						if (p_ds.Tables["target_cond"].Rows.Count == 1)
						{
							strbiosumcondid = Convert.ToString(p_ds.Tables["target_cond"].Rows[0]["biosum_cond_id"]);
						}
						else
						{
							MessageBox.Show("Fatal Error!!! Could not find cond record for cond id " + strcondid);
							this.m_txtStreamWriter.Write("!!!!Fatal Error!!! Could not find cond record for cond id " + strcondid + "\n");
							this.m_intError=-1;
							break;
						}
						p_ds.Tables["target_cond"].Clear();

						strDbFields = "INSERT INTO target_tree_vol_val_by_species_diam_groups (biosum_cond_id";
						strValues=" VALUES ('" + strbiosumcondid + "'";

						if (strrx.Trim().Length > 0)
						{
							strDbFields += ",rx";
							strValues += ",'" + strrx + "'";
						}

					
						if (strspeciesgroup.Trim().Length > 0)
						{
							strDbFields += ",species_group";
							strValues += "," + strspeciesgroup;
						}
						if (strdiamid.Trim().Length > 0)
						{
							strDbFields += ",diam_group";
							strValues += "," + strdiamid;
						}
						if (strbiomassvolcf.Trim().Length > 0)
						{
							strDbFields += ",biomass_vol_cf";
							strValues += "," + strbiomassvolcf;
						}
						if (strbiomasswtgt.Trim().Length > 0)
						{
							strDbFields += ",biomass_wt_gt";
							strValues += "," + strbiomasswtgt;
						}
						if (strmerchvolcf.Trim().Length > 0)
						{
							strDbFields += ",merch_vol_cf";
							strValues += "," + strmerchvolcf;
						}
						if (strmerchwtgt.Trim().Length > 0)
						{
							strDbFields += ",merch_wt_gt";
							strValues += "," + strmerchwtgt;
						}
						if (strmerchvaldpa.Trim().Length > 0)
						{
							strDbFields += ",merch_val_dpa";
							strValues += "," + strmerchvaldpa;
						}
						
						strDbFields += ")";
						strValues += ");";
						strSQL = strDbFields + strValues;

						this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strSQL);
						if (this.m_ado.m_intError < 0)
						{
							this.m_txtStreamWriter.Write("Data Error: insert sql error for sql statement " + strSQL + "\n");
						}
						System.Windows.Forms.Application.DoEvents();
						if (this.m_frmTherm.AbortProcess == true) break;
					}
				}
				p_OleDbCommand = null;
				p_ds.Clear();
				p_da.Dispose();
				p_ds = null;
				p_da = null;
				this.m_txtStreamWriter.Write("tree vol val Records Processed: " + intCurrRec.ToString() + "\n");

		}
		private void convert_haulcost()
		{
			System.Data.OleDb.OleDbDataAdapter p_da;
			System.Data.DataSet p_ds;
			System.Data.OleDb.OleDbCommand p_OleDbCommand;
			p_da = new System.Data.OleDb.OleDbDataAdapter();
			p_ds = new DataSet("tempds");
			this.m_intError=0;
			this.m_strError = "";
			string strSQL="";
			int intCurrRec=0;

			string strpsiteid,strname,strtypecd;
			string strtypecddef,strmaterialcd,strmaterialcddef;
			string strexists_yn,strkeeptimeslayer_yn;
			
			string strDbFields="";
			string strValues="";
			
			this.m_intError=0;
			this.m_txtStreamWriter.Write("\n\nProcessing Site Data\n");
			this.m_txtStreamWriter.Write("--------------------------\n");

			try
			{
				strSQL = "DELETE * FROM TARGET_PSITE";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, strSQL);
				strSQL = "SELECT * FROM SOURCE_PSITE";
				m_frmTherm.AbortProcess = false;
				m_frmTherm.lblMsg.Text = "Converting Processing Site Data";
				m_frmTherm.lblMsg.Visible = true;
				m_frmTherm.lblMsg.Refresh();
				this.m_frmTherm.progressBar1.Maximum = Convert.ToInt32(this.m_ado.getRecordCount("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.strTempMDBFile + ";User Id=admin;Password=;","select count(*) from source_psite","source_psite"));
				this.m_txtStreamWriter.Write("source processing site record count=" + Convert.ToString(this.m_frmTherm.progressBar1.Maximum) + "\n");
				p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
				p_OleDbCommand.CommandText = strSQL;
				p_da.SelectCommand = p_OleDbCommand;
				p_da.Fill(p_ds, "source_psite");
			}
			catch (Exception caught)
			{
				MessageBox.Show("Abort Process: Error in psite conversion with error " + caught.Message);
				this.m_intError=-1;
				p_OleDbCommand = null;
				p_ds = null;
				p_da = null;
				return;
			}
			if (this.m_intError==0)
			{
				for (intCurrRec = 0; intCurrRec <= p_ds.Tables["source_psite"].Rows.Count-1;intCurrRec++)
				{
					strpsiteid="";
					strname="";
					strtypecd="";
					strtypecddef="";
					strmaterialcd="";
					strmaterialcddef="";
					strexists_yn="";
					strkeeptimeslayer_yn="";

					this.m_frmTherm.Increment(intCurrRec + 1);
					strpsiteid = Convert.ToString(p_ds.Tables["source_psite"].Rows[intCurrRec]["psite"]);
					strname = Convert.ToString(p_ds.Tables["source_psite"].Rows[intCurrRec]["name"]);

					strDbFields = "INSERT INTO target_psite (psite_id";
					strValues=" VALUES ('" + strpsiteid + "'";

					if (strname.Trim().Length > 0)
					{
						strDbFields += ",name";
						strValues += ",'" + strname + "'";
					}
					strDbFields += ")";
					strValues += ");";
					strSQL = strDbFields + strValues;

					this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strSQL);
					if (this.m_ado.m_intError < 0)
					{
						this.m_txtStreamWriter.Write("Data Error: insert sql error for sql statement " + strSQL + "\n");
					}
					System.Windows.Forms.Application.DoEvents();
					if (this.m_frmTherm.AbortProcess == true) break;
				}
			}
			p_OleDbCommand = null;
			p_ds.Clear();
			p_da.Dispose();
			p_ds = null;
			p_da = null;
			this.m_txtStreamWriter.Write("psite records Processed: " + intCurrRec.ToString() + "\n");


		}
		private void create_tree()
		{
			string strSQL="";
			string stridbcondid="";
			string stridballtreeid="";
			string strBiosumCondId="";
			string strunitcd="";
			string strcountycd="";
			string strplot="";
			string strsubp="";
			string strtree="";
			string strcondid="";
			string strstatuscd="";
			string strspcd="";
			string strspgrpcd="";
			string strdia="";
			string strdiahtcd="";
			string strht="";
			string strhtcd="";
			string stractualht="";
			string strtreeclcd="";
			string strcr="";
			string strcclcd="";
			string strcull="";
			string strdecaycd="";
			string strstocking="";
			string strtpacurr="";
			string strvolcfnet="";
			string strvolcfgrs="";
			string strvolcsnet="";
			string strvolcsgrs="";
			string strvolbfnet="";
			string strvolbfgrs="";
			string strdrybiot="";
			string strdrybiom="";
			string strdiacheck="";
			string strbhage="";
			string strcullbf="";
			string strcullsf="";
			string strtotage="";
			string strmistclcdpnwrs="";
			string stragentcd="";
			string strdamtyp1="";
			string strdamsev1="";
			string strdamtyp2="";
			string strdamsev2="";
			string stridbdmgagent3cd="";
			string stridbseverity3cd="";

			
			

			System.Data.OleDb.OleDbDataAdapter p_da;
			System.Data.DataSet p_ds;
			System.Data.OleDb.OleDbCommand p_OleDbCommand;
			p_da = new System.Data.OleDb.OleDbDataAdapter();
			p_ds = new DataSet("tempds");
			this.m_intError=0;
			this.m_strError = "";
			
			int intCurrRec=0;

			try
			{
				strSQL = "DELETE FROM TARGET_TREE";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, strSQL);
				if (this.m_ado.m_intError==0)
				{
					strSQL = "SELECT * FROM TARGET_COND";
					m_frmTherm.AbortProcess = false;
					m_frmTherm.lblMsg.Text = "Converting Tree Data";
					m_frmTherm.lblMsg.Visible = true;
					m_frmTherm.lblMsg.Refresh();
					this.m_frmTherm.progressBar1.Maximum = Convert.ToInt32(this.m_ado.getRecordCount("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.strTempMDBFile + ";User Id=admin;Password=;","select count(*) from target_cond","target_cond"));
					this.m_txtStreamWriter.Write("source cond table record count=" + Convert.ToString(this.m_frmTherm.progressBar1.Maximum) + "\n");
					p_OleDbCommand = this.m_TempMDBFileConn.CreateCommand();
					p_OleDbCommand.CommandText = strSQL;
					p_da.SelectCommand = p_OleDbCommand;
					p_da.Fill(p_ds, "target_cond");
					for (intCurrRec = 0; intCurrRec <= p_ds.Tables["target_cond"].Rows.Count-1;intCurrRec++)
					{
						stridbcondid="";
						stridballtreeid="";
						strBiosumCondId="";
						strunitcd="";
						strcountycd="";
						strplot="";
						strsubp="";
						strtree="";
						strcondid="";
						strstatuscd="";
						strspcd="";
						strspgrpcd="";
						strdia="";
						strdiahtcd="";
						strht="";
						strhtcd="";
						stractualht="";
						strtreeclcd="";
						strcr="";
						strcclcd="";
						strcull="";
						strdecaycd="";
						strstocking="";
						strtpacurr="";
						strvolcfnet="";
						strvolcfgrs="";
						strvolcsnet="";
						strvolcsgrs="";
						strvolbfnet="";
						strvolbfgrs="";
						strdrybiot="";
						strdrybiom="";
						strdiacheck="";
						strbhage="";
						strcullbf="";
						strcullsf="";
						strtotage="";
						strmistclcdpnwrs="";
						stragentcd="";
						strdamtyp1="";
						strdamsev1="";
						strdamtyp2="";
						strdamsev2="";
						stridbdmgagent3cd="";
						stridbseverity3cd="";
						stridballtreeid="";

						if (this.m_ado.m_intError < 0)
						{
							this.m_txtStreamWriter.Write("Data Error: insert sql error for sql statement " + strSQL + "\n");
						}
						System.Windows.Forms.Application.DoEvents();
						if (this.m_frmTherm.AbortProcess == true) break;
					}
					
				}
			}
			catch (Exception e)
			{
				MessageBox.Show(e.Message);
			}


			
			
			
			
			
		}
		private void ThermCancel(object sender, System.EventArgs e)
		{
			string strMsg = "Do you wish to cancel the data conversion process (Y/N)?";
			DialogResult result = MessageBox.Show(strMsg,"Cancel Process", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			switch (result) 
			{
				case DialogResult.Yes:
					this.m_frmTherm.AbortProcess = true;
					this.m_frmTherm.Hide();
					return;
				case DialogResult.No:
					return;
			}                
		}
	}
}