using System;
using System.Threading;
using System.IO;
using System.Globalization;
using System.Text;
using System.Data;
using System.Data.Common;
using System.Collections;
using System.Windows.Forms;
using Microsoft.Office.Interop.Access.Dao;
using OLE=System.Data.OleDb;

namespace FIA_Biosum_Manager
{
   /// <summary>A class that regulates the processes of post FVS and pre FRCS</summary>
   public class run_PostFvsForeFrcs
   {
      #region Private Variables
      private const string cStrFvsSaveFN=@"\Save";
      private const string cStrRxExtFilter=@"*rx*.mdb";
      private const string cStrRxFilter="rx='{0}'";
      private const string cStrStandIdFilter="StandID={0}";
      private const string cStrBiosumCondIdFilter="biosum_cond_id='{0}'";
      private const string cStrSelectTable=@"SELECT * FROM {0} WHERE {1} = {2}";
      private const string strPathProject=@"C:\FiaBiosumProjects\fia_biosum\orca_converted_data";      //TODO: Use Datasource
      private string strPathMasterDbFile=strPathProject+@"\db\master.mdb";                                                        //TODO: Use Datasource
      private string strPathFvsDbFile=strPathProject+@"\fvs\db\out\fvsout_rx{0}.mdb";                                //TODO: Use Datasource
      private string strPathProcDbFile=strPathProject+@"\processor\db\fvs_out_processor_in.mdb";      //TODO: Use Datasource
      private string strDbConx=@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};User Id=admin;Password=;";
      private string strAllegedCondId;
      private ArrayList aryLRxSelected;
      private Hashtable htUserFvsoutSet;
      private DataSet dsWorkDataSet;

       //ldp added

	   private string m_strProjDir;
	   private FIA_Biosum_Manager.Datasource m_DataSource;
	   public int m_intError=0;
	   private string m_strCondTable;
	   private string m_strFFETable;
	   private string m_strProcessorInTreeListTable;
	   private string m_strRxTable;
	   private string m_strTempMDBFile;
	   private ado_data_access m_ado;
	   private dao_data_access m_dao;
	   private string m_strConn;

      #endregion

      #region Public Properties
      public string PathProject{get{return strPathProject;}}
      public string ConditionIdString{get{return strAllegedCondId;}set{strAllegedCondId=value;}}
      public ArrayList WfSelectedRxList{get{return aryLRxSelected;}}
      public DataSet WorkDataSet{get{return dsWorkDataSet;}}
      #endregion

      #region Public Methods
      //TOP LEVEL call - after setting the selection of RX files...
      public void WfStart()
      {
         Thread t=new Thread(new ThreadStart(Wf_ReSelectedRx));
         t.Start();
         t.Priority=ThreadPriority.Highest;         
         t.Join();
      }

      //assumed this.WfSelectedRxList is SET
      public void Wf_ReSelectedRx()
      {         
         for(short x=0; x<this.aryLRxSelected.Count; x++)
            this.WfAll(this.aryLRxSelected[x].ToString());
      }


      //1) UI - Save Validated RX DB files -- If not present then and not populated?      
      public void WfAll(string strFilePathRx)
      {         
         
      }

 
      public void WfFinalize()
      {
         this.DbRecsFinalizeFvsTree("");
         this.CleanUp();
      }
      #endregion

      #region Private Methods
      //get the Cond, Fvs_tree and ffe tables from master, passing IN: ((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory;
      private DataRow[] DbRecsDataForFVS(string strProjDir) 
      {         
         DataTable dtFFE=new DataTable("FFE");
         DataTable dtCond=new DataTable("COND");

         ado_core_tables adoCore=new ado_core_tables();
         adoCore.m_OleDbConnection.ConnectionString=string.Format(this.strDbConx, strProjDir);
         adoCore.m_OleDbConnection.Open();
         
         adoCore.m_daFFE=new System.Data.OleDb.OleDbDataAdapter("SELECT * FROM FFE", adoCore.m_OleDbConnection.ConnectionString);
         adoCore.m_daCond=new System.Data.OleDb.OleDbDataAdapter("SELECT  * FROM COND", adoCore.m_OleDbConnection.ConnectionString);
         
         adoCore.m_daFFE.Fill(dtFFE);
         adoCore.m_daCond.Fill(dtCond);
         this.WorkDataSet.Tables.Add(dtFFE);
         this.WorkDataSet.Tables.Add(dtCond);
         
      
         return this.WorkDataSet.Tables["FFE"].Select("", "");
      }

      
      //3.5.0.0 & 3.6.0.0
      //TODO: Change back to Private and reduce the CONDID to one joined column
      public int DbRecsGetCondTabPreValues(string strCondIdMasterCondId, string strCondId_PotFireStandId, string strCurRxFilePath)
      {         
         int intRet=0;
         string strSqlCond=@"SELECT * FROM COND WHERE BIOSUM_COND_ID = '{0}'";         
         string strSqlFvsPotFire=@"SELECT * FROM FVS_POTFIRE WHERE STANDID = '{0}'";

         //TODO: Missing not_calc_yn,  & what happens if the row count dn match - the row Cond ID must be less?
         string[] strSqlCondFilter=new string[12]{"pre_Tot_Flame_Mod", "pre_Tot_Flame_Sev", "pre_Fire_Type_Severe", "pre_Fire_Type_Mod", "pre_Torch_Index", "pre_Crown_Index", "pre_Canopy_Ht", "pre_Canopy_Density", "pre_Mortality_BA_Severe", "pre_Mortality_BA_Mod", "pre_Mortality_VOL_Severe", "pre_Mortality_VOL_Mod"};

         string[] strSqlFvsPfFilter=new string[12]{"Flame_Len_Mod", "Flame_Len_Severe", "Fire_Type_Severe", "Fire_Type_Mod", "Torch_Index", "Crown_Index", "Canopy_Ht", "Canopy_Density", "Mortality_BA_Severe", "Mortality_BA_Mod", "Mortality_VOL_Severe", "Mortality_VOL_Mod"};                      


         DataTable dtCond=new DataTable("COND");
         DataTable dtFvsPF=new DataTable("FVS_PF");         
         try
         {
            ado_core_tables adoCore=new ado_core_tables();
            adoCore.m_OleDbConnection=new System.Data.OleDb.OleDbConnection(string.Format(this.strDbConx, this.strPathMasterDbFile));         
            adoCore.m_OleDbConnection.Open();
            adoCore.m_daCond=new System.Data.OleDb.OleDbDataAdapter(string.Format(strSqlCond, strCondIdMasterCondId), adoCore.m_OleDbConnection.ConnectionString);            
            adoCore.m_daCond.Fill(dtCond);

            ado_core_tables adoFvs=new ado_core_tables();
            adoFvs.m_OleDbConnection=new System.Data.OleDb.OleDbConnection(string.Format(this.strDbConx, strCurRxFilePath));
            adoFvs.m_OleDbConnection.Open();
            adoFvs.m_daCond=new System.Data.OleDb.OleDbDataAdapter(string.Format(strSqlFvsPotFire, strCondId_PotFireStandId), adoFvs.m_OleDbConnection.ConnectionString);
            OLE.OleDbCommandBuilder oleCmdB=new System.Data.OleDb.OleDbCommandBuilder(adoFvs.m_daCond);
            adoFvs.m_daCond.Fill(dtFvsPF);
                  
            if(dtCond.Rows.Count <= dtFvsPF.Rows.Count)
            {
               for(short i=0; i<dtCond.Rows.Count; i++)
               {
                  dtCond.Rows[i].BeginEdit();
                  for(short x=0; x<strSqlCondFilter.Length; x++)
                  {
                     dtCond.Rows[i][strSqlCondFilter[x]]=dtFvsPF.Rows[i][strSqlFvsPfFilter[x]];
                     intRet++;
                  }
                  dtCond.Rows[i].EndEdit();
               }
               adoFvs.m_daCond.Update(dtCond);
               dtCond.AcceptChanges();               
            }

         }
         catch(Exception e)
         {            
            intRet=0;
            throw e;
         }
         return intRet;
      }


      //TODO: change back to Private
      public int DbRecsFinalizeFvsTree(string strRxFilePath)
      {
         int intRet=0;     
         int intExe=0;
         string strRxId=this.UtilHelperGetRxString(strRxFilePath);
         string strSqlSrc="SELECT * FROM FVS_CUTLIST";             
         string strSqlTgt="SELECT * FROM FVS_TREE";
                                                                                                                     //N,N,T
         string[] strSqlSrcFilterFvsCutlist=new string[27]{"Id", "StandID", "getRxId",  "Year", "PrdLen", "TreeId", "Species", "TreeVal", "SSCD", "PtIndex", "TPA", "MortPA", "DBH", "DG", "Ht", "HtG", "PctCr", "CrWidth", "MistCD", "BAPctile", "PtBAL", "TCuFt", "MCuFt", "BdFt", "MDefect", "BDefect", "TruncHt"};                                                                         //N,T,T
         string[] strSqlTgtFilterFvsTree=new string[27]{"Id", "biosum_cond_id", "rx", "Year", "PrdLen", "Tid", "Species", "TreeValue", "SSCD", "PtIndex", "TPA", "MortPA", "DBH", "DG", "Ht", "HtG", "PctCr", "CrWidth", "MistCD", "BAPctile", "PtBAL", "TCuFt", "MCuFt", "BdFt", "MDefect", "BDefect", "TruncHt"};
         
         DataTable dtTgtFvsTree=null;
         DataTable dtSrcFvsCutList=null;
         OLE.OleDbConnection oleConxTgt=null;
         OLE.OleDbDataAdapter oleAdaptTgt=null;
         OLE.OleDbCommandBuilder oleCmdBTgt=null;
         OLE.OleDbCommand oCmd=null;
         OLE.OleDbConnection oleConxSrc=null;
         OLE.OleDbDataAdapter oleAdaptSrc=null;         
         try
         {
            oleConxTgt=new System.Data.OleDb.OleDbConnection(string.Format(this.strDbConx, this.strPathProcDbFile));
            oleConxTgt.Open();
            oleAdaptTgt=new System.Data.OleDb.OleDbDataAdapter(strSqlTgt, oleConxTgt);
            dtTgtFvsTree=new DataTable("FVS_Tree");
            oleCmdBTgt=new System.Data.OleDb.OleDbCommandBuilder(oleAdaptTgt);
            oleAdaptTgt.Fill(dtTgtFvsTree);                        

            if(dtTgtFvsTree.Rows.Count>0)
            {
               DataRow[] drDel=dtTgtFvsTree.Select(string.Format(cStrRxFilter, strRxId));
               if(drDel.Length>0)
               {
                  foreach(DataRow dr in drDel)
                  {
                     dr.Delete(); //delete old Records base on RX ID
                  }
                  oleAdaptTgt.Update(dtTgtFvsTree);
                  dtTgtFvsTree.AcceptChanges();
               }
            }
            oleConxSrc=new System.Data.OleDb.OleDbConnection(string.Format(this.strDbConx, string.Format(this.strPathFvsDbFile, strRxId)));
            oleConxSrc.Open();
            oleAdaptSrc=new System.Data.OleDb.OleDbDataAdapter(strSqlSrc, oleConxSrc);
            dtSrcFvsCutList=new DataTable("FVS_CUTLIST");
            oleAdaptSrc.Fill(dtSrcFvsCutList);

            if(dtSrcFvsCutList.Rows.Count>1)
            {                                             
               dtTgtFvsTree.BeginLoadData();
               foreach(DataRow dr in dtSrcFvsCutList.Rows)
               {                                    
                  dtTgtFvsTree.LoadDataRow(dr.ItemArray, false);                  
                  DataRow drCur=dtTgtFvsTree.Rows[intRet];
                  drCur.BeginEdit();
                  drCur["rx"]=strRxId;
                  drCur["Tid"]=dr["TreeId"].ToString().Trim();     //Data Quality Issue
                  drCur.EndEdit();    
                  intRet++;    
               }
               dtTgtFvsTree.EndLoadData();               
               dtTgtFvsTree.AcceptChanges();
               string[] strArySql=this.DbHelperStrSqlInserts(dtTgtFvsTree);                 //See Documentation
               oCmd=oleConxTgt.CreateCommand();
               for(short x=0; x<strArySql.Length; x++)
               {
                  oCmd.CommandText=strArySql[x].Trim();
                  oCmd.CommandType=CommandType.Text;
                  intExe=oCmd.ExecuteNonQuery();
               }
            }            
         }
         catch(OLE.OleDbException oe)
         {
            throw oe;
         }
         catch(Exception e)
         {
            throw e;
         }
         finally
         {
            if(oleConxSrc.State==ConnectionState.Open)
               oleConxSrc.Close();
            if(oleConxTgt.State==ConnectionState.Open)
               oleConxTgt.Close();
            oleConxSrc=null;
            oleConxTgt=null;
         }
         return intExe;
      }

      //TODO: provide the UI interface CheckBox to overwrite existing differing Diam and Species Group from Biosum and FVS
      //Otherwise, make a copy? and or not do anything?
      //TODO: Change back to Private
      public void DbRecsTdgTsg2ProcDbCopy(bool overwrite)
      {
         if(!overwrite)
            return;          //can't do anything now if user don't want replace table yet.

         //bool blnRowDiff=false;

         //15-MAY-2017: These tables were renamed and moved but this code is obsolete so it was not updated
         string strSqlTdg="SELECT * FROM TREE_DIAM_GROUPS";
         string strSqlTsg="SELECT * FROM TREE_SPECIES_GROUPS";
         string strSqlDelTdg="DELETE FROM TREE_DIAM_GROUPS";
         string strSqlDelTsg="DELETE FROM TREE_SPECIES_GROUPS";
         
         DataTable dtTdg=new DataTable("TDG");
         DataTable dtTsg=new DataTable("TSG");;
         DataTable dtTdgTgt=new DataTable("TDGTGT");
         DataTable dtTsgTgt=new DataTable("TSGTGT");
         DataSet dsMaster=new DataSet("FVS_TREE");
         dsMaster.Tables.Add(dtTdg);
         dsMaster.Tables.Add(dtTsg);

         OLE.OleDbConnection oleConx2Master=null;
         OLE.OleDbDataAdapter oleAdaptTdg=null;
         OLE.OleDbDataAdapter oleAdaptTsg=null;         
         OLE.OleDbConnection oleConx2Proc=null;
         OLE.OleDbDataAdapter oleAdaptTdgTgt=null;
         OLE.OleDbDataAdapter oleAdaptTsgTgt=null;
         OLE.OleDbCommandBuilder oleCmdBTdgTgt=null;
         OLE.OleDbCommandBuilder oleCmdBTsgTgt=null;

         try
         {
            oleConx2Master=new System.Data.OleDb.OleDbConnection(string.Format(this.strDbConx, this.strPathMasterDbFile));
            oleConx2Master.Open();
            oleAdaptTdg=new System.Data.OleDb.OleDbDataAdapter(strSqlTdg, oleConx2Master);            
            oleAdaptTdg.Fill(dtTdg);
            oleAdaptTsg=new System.Data.OleDb.OleDbDataAdapter(strSqlTsg, oleConx2Master);            
            oleAdaptTsg.Fill(dtTsg);
            
            //TODO: Made assumption that the table structures exist ... need to create the Table 
            //If 2 tables exists in strPathProcDbFile, then delete records and import them from here.      
            oleConx2Proc=new System.Data.OleDb.OleDbConnection(string.Format(this.strDbConx, this.strPathProcDbFile));
            oleConx2Proc.Open();
            oleAdaptTdgTgt=new System.Data.OleDb.OleDbDataAdapter(strSqlTdg, oleConx2Proc);
            oleAdaptTsgTgt=new System.Data.OleDb.OleDbDataAdapter(strSqlTsg, oleConx2Proc);
            oleAdaptTdgTgt.Fill(dtTdgTgt);
            oleAdaptTsgTgt.Fill(dtTsgTgt);
            oleCmdBTdgTgt=new System.Data.OleDb.OleDbCommandBuilder(oleAdaptTdgTgt);
            oleCmdBTsgTgt=new System.Data.OleDb.OleDbCommandBuilder(oleAdaptTsgTgt);            
            
            int tdgMatch=0, tsgMatch=0;
            foreach(DataRow dr in dtTdg.Rows)
            {
               foreach(DataRow dr2 in dtTdgTgt.Rows)
               {
                  if(dr.ItemArray.ToString().CompareTo(dr2.ItemArray.ToString())==0)
                     tdgMatch++;
               }
            }
            foreach(DataRow dr in dtTsg.Rows)
            {
               foreach(DataRow dr2 in dtTsgTgt.Rows)
               {
                  if(dr.ItemArray.ToString().CompareTo(dr2.ItemArray.ToString())==0)
                     tsgMatch++;
               }
            }         
            //flag value compare difference in both row and value
            if(tdgMatch!=dtTdg.Rows.Count||tsgMatch!=dtTsg.Rows.Count||dtTdg.Rows.Count!=dtTdgTgt.Rows.Count||dtTsg.Rows.Count!=dtTsgTgt.Rows.Count)
              // blnRowDiff=true;

            if(dtTdgTgt.Rows.Count>0 && overwrite)
            {
               OLE.OleDbCommand oCmd=oleConx2Proc.CreateCommand();
               oCmd.CommandText=strSqlDelTdg;
               oCmd.ExecuteNonQuery();
               dtTdgTgt.Rows.Clear();
               dtTdgTgt.AcceptChanges();
            } 
            if(dtTsgTgt.Rows.Count>0 && overwrite)
            {
               OLE.OleDbCommand oCmd=oleConx2Proc.CreateCommand();
               oCmd.CommandText=strSqlDelTsg;
               oCmd.ExecuteNonQuery();
               dtTsgTgt.Rows.Clear();
               dtTsgTgt.AcceptChanges();
            }
                           
            dtTdgTgt.BeginLoadData();
            dtTsgTgt.BeginLoadData();
            foreach(DataRow dr in dtTdg.Rows)
            {
              dtTdgTgt.LoadDataRow(dr.ItemArray, false);                
            }
            foreach(DataRow dr in dtTsg.Rows)
            { 
               dtTsgTgt.LoadDataRow(dr.ItemArray, false);               
            }
            dtTdgTgt.EndLoadData();
            dtTsgTgt.EndLoadData();                        
            oleAdaptTdgTgt.Update(dtTdgTgt);
            dtTdgTgt.AcceptChanges();
            oleAdaptTsgTgt.Update(dtTsgTgt);
            dtTsgTgt.AcceptChanges();
         }
         catch(OLE.OleDbException oe)
         {
            throw oe;
         }
         finally
         {
            if(oleConx2Master.State==ConnectionState.Open)
               oleConx2Master.Close();            
            oleConx2Master=null;
            oleAdaptTdg=null;
            oleAdaptTsg=null;
            if(oleConx2Proc.State==ConnectionState.Open)
               oleConx2Proc.Close();            
            oleAdaptTdgTgt=null;
            oleAdaptTsgTgt=null;            
         }
         
      }

      //1) Delete the recs from FFE with related RX ID
      //2) Append into FFE from FVS_Potfire
      public int DbRecsUpdateFfeFromPotFire(string strRxFilePath)
      {
         int intRet=0;
         string strRxId=this.UtilHelperGetRxString(strRxFilePath);
         string strSqlSrcPotFire="SELECT * FROM FVS_PotFire";
         string strSqlTgtFfe="SELECT * FROM FFE WHERE biosum_cond_id = '{0}'";                                                                        // TODO: NOT_CALC_YN VALUE?
         string[] strSqlSrcFilterPotFire=new string[15]{"StandID", "CaseID", " ", "Flame_Len_Severe", "Flame_Len_Mod", "Fire_Type_Severe", "Fire_Type_Mod", "Torch_Index", "Crown_Index", "Canopy_Ht", "Canopy_Density", "Mortality_BA_Severe", "Mortality_BA_Mod", "Mortality_VOL_Severe", "Mortality_VOL_Mod"};
         string[] strSqlTgtFilterFfe=new string[15]{"biosum_cond_id", "rx", "post_not_calc_yn", "post_tot_flame_sev", "post_tot_flame_mod", "post_fire_type_sev", "post_fire_type_mod", "post_torch_index", "post_crown_index", "post_canopy_ht",  "post_canopy_density", "post_mortality_ba_sev", "post_mortality_ba_mod", "post_mortality_vol_sev", "post_mortality_vol_mod"};
         
         DataTable dtSrcPotFire=null;
         DataTable dtTgtFfe=null;
         OLE.OleDbConnection oleConxSrc=null;
         OLE.OleDbDataAdapter oleAdaptSrc=null;
         OLE.OleDbConnection oleConxTgt2=null;
         OLE.OleDbDataAdapter oleAdaptTgt2=null;
         OLE.OleDbCommandBuilder oleCmdBTgt2=null;
         try
         {          
            oleConxSrc=new System.Data.OleDb.OleDbConnection(string.Format(this.strDbConx, string.Format(this.strPathFvsDbFile, strRxId)));
            oleConxSrc.Open();
            oleAdaptSrc=new System.Data.OleDb.OleDbDataAdapter(string.Format(strSqlSrcPotFire, strRxId), oleConxSrc);
            dtSrcPotFire=new DataTable("FVSPOTFIRE");
            oleAdaptSrc.Fill(dtSrcPotFire);

            oleConxTgt2=new System.Data.OleDb.OleDbConnection(string.Format(this.strDbConx, this.strPathMasterDbFile));
            oleConxTgt2.Open();
            oleAdaptTgt2=new System.Data.OleDb.OleDbDataAdapter(strSqlTgtFfe, oleConxTgt2);
            dtTgtFfe=new DataTable("FFE");
            oleAdaptTgt2.Fill(dtTgtFfe);
            oleCmdBTgt2=new System.Data.OleDb.OleDbCommandBuilder(oleAdaptTgt2);

            if(dtTgtFfe.Rows.Count>0)
            {
               DataRow[] drs=dtTgtFfe.Select(string.Format(cStrRxFilter, strRxId));
               foreach(DataRow dr in drs)
                  dr.Delete();
               oleAdaptTgt2.Update(dtTgtFfe);
               dtTgtFfe.AcceptChanges();
            }


            if(dtSrcPotFire.Rows.Count>1)
            {
               dtTgtFfe.BeginLoadData();
               for(short i=0; i<dtSrcPotFire.Rows.Count; i++)
               {                 
                  dtTgtFfe.LoadDataRow(dtSrcPotFire.Rows[i].ItemArray, false);
                  dtTgtFfe.Rows[i]["rx"]=strRxId;                                     
                  intRet++;
               }
               dtTgtFfe.EndLoadData();
            }
            oleAdaptTgt2.Update(dtTgtFfe);
            dtTgtFfe.AcceptChanges();
         }
         catch(Exception e)
         {
            throw e;
         }
         finally
         {
            if(oleConxTgt2.State==ConnectionState.Open)
               oleConxTgt2.Close();
            oleConxTgt2=null;
         }
         return 0;
      }

 
      //Get the Select files and check to see how many records are in the selected Verified files, TODO: otherwise?
      //Check for null when calling
      private Hashtable DbRecsCountVerify(ArrayList aryOfExistingFvsRxMdbFiles)
      {
         Hashtable htTmp=null;
         string strSql="SELECT COUNT(*) FROM FVS_TREELIST";
         if(aryOfExistingFvsRxMdbFiles.Count<1)
            return htTmp;
         else
         {
            htTmp=new Hashtable();
            ado_data_access adoD=new ado_data_access();
            for(short x=0; x<aryOfExistingFvsRxMdbFiles.Count; x++)
            {
               htTmp.Add(aryOfExistingFvsRxMdbFiles[x].ToString(), Convert.ToInt32(adoD.getRecordCount(string.Format(this.strDbConx, string.Format(this.strPathFvsDbFile,aryOfExistingFvsRxMdbFiles[x].ToString())), strSql, "FVS_TREELIST")));
            }
         }
         return htTmp;
      }

      //Get rx listing from project and make sure all FSOUT_rx* is present, TODO: otherwise?
      private ArrayList DbRxVerify()
      {
         ArrayList aryL=new ArrayList();
         ado_data_access  p_ado = new ado_data_access();
         
         string strSQL =@"SELECT * FROM RX;";
         try
         {
            p_ado.SqlQueryReader(string.Format(this.strDbConx, this.strPathMasterDbFile), strSQL);
            if (p_ado.m_intError==0)
            {
               if (p_ado.m_OleDbDataReader.HasRows == true)
               {
                  while(p_ado.m_OleDbDataReader.Read())
                     aryL.Add(p_ado.m_OleDbDataReader["RX"].ToString().Trim());
               }
               else 
               {

                  //TODO: Log to file
               }
            }
         }
         catch(OLE.OleDbException ode)
         {            
            p_ado.m_OleDbDataReader.Close();
            p_ado.m_OleDbConnection.Close();			
            throw ode;
         }
         finally
         {
            p_ado.m_OleDbDataReader.Close();
            p_ado.m_OleDbConnection.Close();			 
            p_ado.m_OleDbDataReader=null;
            p_ado.m_OleDbConnection = null;
            p_ado=null;
         }
         return aryL;
      }


      private string UtilHelperGetRxString(string inStr)
      {
        return inStr.Substring(inStr.Length-5, 1);
      }

      private bool UtilHelperIsStringType(string strType)
      {
         switch (strType.Trim())
         {
            case "System.Single":
               return false;
            case "System.Double":
               return false;
            case "System.Decimal":
               return false;
            case "System.String":
               return true;
            case "System.Int16":
               return false;
            case "System.Char":
               return true;
            case "System.Int32":
               return false;
            case "System.DateTime":
               return false;
            case "System.DayOfWeek":
               return false;
            case "System.Int64":
               return false;
            case "System.Byte":
               return false;
            case "System.Boolean":
               return false;
            default:
               return false;
         }
      }

      private string[] DbHelperStrSqlInserts(DataTable dt)
      {         
         StringBuilder sbFinalizer=new StringBuilder(dt.Rows.Count);
         StringBuilder sbColName=null;
         StringBuilder sbColValue=null;
         string[] strAryRet=new string[dt.Rows.Count];
         try
         {
            for(short x=0; x<dt.Rows.Count; x++)
            {
               sbColName=new StringBuilder(dt.Rows[x].ItemArray.Length);
               sbColValue=new StringBuilder(dt.Rows[x].ItemArray.Length);
               for(short y=0; y<dt.Rows[x].ItemArray.Length; y++)
               {
                  //cols 
                  sbColName.AppendFormat("{0},", dt.Rows[x].Table.Columns[y].ToString());
                  //vals
                  if(!this.UtilHelperIsStringType(dt.Rows[x].ItemArray[y].GetType().ToString()))
                     sbColValue.AppendFormat("{0},", dt.Rows[x].ItemArray[y]);
                  else
                     sbColValue.AppendFormat("'{0}',", dt.Rows[x].ItemArray[y]);            
               }               
               //final
               if(sbColName.Length>2)sbColName.Remove(sbColName.Length-1, 1);
               if(sbColValue.Length>2)sbColValue.Remove(sbColValue.Length-1, 1);
               sbFinalizer.AppendFormat("INSERT INTO {0} ({1}) VALUES ({2})", dt.TableName, sbColName.ToString(), sbColValue.ToString());                              
               strAryRet[x]=sbFinalizer.ToString();
               sbFinalizer.Remove(0, sbFinalizer.Length);
            }
         }
         catch(Exception e)
         {
            throw e;
         }         
         return strAryRet;
      }

      //returns the missing RX Files
      public Hashtable FioIsRxExist(string strPath)
      {
         StringBuilder strb=new StringBuilder();
         strb.Append(strPath).Append(@"\fvsout_rx");         
                  
         ArrayList aryL=this.DbRxVerify();
         Hashtable htFound=new Hashtable(aryL.Count);
         try
         {
            for(short x=0; x<aryL.Count; x++)
            {               
               strb.Append(aryL[x].ToString());
               strb.Append(@".mdb");
               if(File.Exists(strb.ToString()))
               {
                  htFound.Add(this.FioGetAbsFileName(strb.ToString()), strb.ToString());
                  strb.Remove(strb.Length-5, 5);
               }
               else
               {
                  strb.Remove(strb.Length-5, 5);
                  continue;
               }
            }
         }
         catch(IOException ie)
         {                                 
            throw ie;
         }
         return htFound;
      }     

      //save the copies of the FVS MDB files Call: ((frmMain)this.ParentForm).frmProject.uc_project1.txtRootDirectory.Text.Trim() 
      private void FioArchiveFiles(string[] strPathFiles, string strPathBase)
      {         
         foreach(string fileName in strPathFiles)
         {
            if(File.Exists(fileName) && Directory.Exists(strPathBase))
               File.Copy(fileName, string.Format("{0}{1}{2}_{3}",strPathBase, cStrFvsSaveFN, DateTime.Now.ToString("yyyy-MM-dd"), this.FioGetAbsFileName(fileName)), true);            
         }
      }

      private string FioGetAbsFileName(string strSrc)
      {
         string[] strAry=strSrc.Split("\\".ToCharArray());
         ArrayList aryL=new ArrayList(strAry);
         return aryL[aryL.Count-1].ToString();
      }

   
      private void CleanUp()
      {

         this.dsWorkDataSet.Clear();
      }

      #endregion

      #region Constructor
      public run_PostFvsForeFrcs(string p_strProjDir)
      {         

		  this.m_strProjDir = p_strProjDir.Trim();
		  m_DataSource = new Datasource();
		  m_DataSource.LoadTableColumnNamesAndDataTypes=false;
		  m_DataSource.LoadTableRecordCount=false;
		  m_DataSource.m_strDataSourceMDBFile = p_strProjDir.Trim() + "\\db\\project.mdb";
		  m_DataSource.m_strDataSourceTableName = "datasource";
		  m_DataSource.m_strScenarioId="";
		  m_DataSource.populate_datasource_array();

		  this.m_strCondTable = this.m_DataSource.getValidDataSourceTableName("CONDITION");

		  this.m_strProcessorInTreeListTable = this.m_DataSource.getValidDataSourceTableName("FVS TREE LIST FOR PROCESSOR");
		  this.m_strRxTable = this.m_DataSource.getValidDataSourceTableName("TREATMENT PRESCRIPTIONS");
		  this.m_strTempMDBFile = this.m_DataSource.CreateMDBAndTableDataSourceLinks();
		  if (this.m_strCondTable.Trim().Length == 0) 
		  {
			  MessageBox.Show("!!Could Not Locate Cond Table!!","Process FVS Out",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			  this.m_intError=-1;
			  return;
		  }
		  if (this.m_strFFETable.Trim().Length == 0)
		  {
			  MessageBox.Show("!!Could Not Locate FFE Table!!","Process FVS Out",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			  this.m_intError=-1;
			  return;
		  }
		  if (this.m_strProcessorInTreeListTable.Trim().Length == 0)
		  {
			  MessageBox.Show("!!Could Not Locate Processor In Tree Table!!","Process FVS Out",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			  this.m_intError=-1;
			  return;
		  }
		  if (this.m_strRxTable.Trim().Length == 0)
		  {
			  MessageBox.Show("!!Could Not Locate Treatment Table!!","Process FVS Out",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			  this.m_intError=-1;
			  return;
		  }

    	  this.m_ado = new ado_data_access();
		  this.m_strConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"","");
		  this.m_dao = new dao_data_access();
		 



         this.htUserFvsoutSet=new Hashtable();
         this.dsWorkDataSet=new DataSet("FVS2FRCS");
         this.aryLRxSelected=new ArrayList();
      }
      #endregion
   }
}
