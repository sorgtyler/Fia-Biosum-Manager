using System;
using FIA_Biosum_Manager;
using System.Windows.Forms;
namespace FIA_Biosum_Travel_Times_Generator
{

	/// <summary>
	/// Summary description for generate_travel_times.
	/// </summary>
	public class generate_travel_times
	{
		FIA_Biosum_Manager.frmMain m_frmMain;
		FIA_Biosum_Manager.frmTherm m_frmTherm;
		public generate_travel_times(FIA_Biosum_Manager.frmMain p_frmMain)
		{
			this.m_frmMain = p_frmMain;

			//
			// TODO: Add constructor logic here
			//
		}
		public void create_travel_times()
		{
			int intRandomNumber;
			string strRandomNumber;
			string strRandomSave;
            string strMDBFile;
			string strConn;
			string strSQL;
			//string strPsiteRailTravel_YN="N";
			int x;
            int y;
			System.Data.OleDb.OleDbConnection p_conn;
            p_conn = new System.Data.OleDb.OleDbConnection();
			
            ado_data_access p_ado = new ado_data_access();
			
			utils p_utils = new utils();
			strMDBFile = this.m_frmMain.frmProject.uc_project1.txtRootDirectory.Text + "\\gis\\db\\gis_travel_times.accdb";
			strConn=p_ado.getMDBConnString(strMDBFile,"admin","");
			p_ado.OpenConnection(strConn,ref p_conn);
            p_ado.SqlNonQuery(p_conn,"delete from travel_time");
			if (p_ado.m_intError==0)
			{
				p_ado.CreateDataSet(p_conn,"select * from plot","plot");
				if (p_ado.m_intError==0)
				{
					p_ado.AddSQLQueryToDataSet(p_conn,ref p_ado.m_OleDbDataAdapter,ref p_ado.m_DataSet,"select * from processing_site where ucase(name) like 'TEST%'","processing_site");
					if (p_ado.m_intError==0)
					{
					
					    this.m_frmTherm = new frmTherm();
						this.m_frmTherm.AbortProcess=false;
						this.m_frmTherm.progressBar1.Minimum = 1;
						this.m_frmTherm.progressBar1.Maximum = 
							p_ado.m_DataSet.Tables["plot"].Rows.Count;
						m_frmTherm.btnCancel.Visible=true;
						m_frmTherm.Show();
						m_frmTherm.Focus();
						m_frmTherm.Text = "Generating Random Travel Times";
						m_frmTherm.Refresh();
						this.m_frmTherm.btnCancel.Click += new System.EventHandler(this.ThermCancel);
						for (x=0;x<=p_ado.m_DataSet.Tables["plot"].Rows.Count-1;x++)
						{
							this.m_frmTherm.Increment(x + 1);
							strRandomSave="";
							for (y=0;y<=p_ado.m_DataSet.Tables["processing_site"].Rows.Count-1;y++)
							{
								
								for (;;)
								{
									intRandomNumber = p_utils.RandomNumber(1,18);
									strRandomNumber = Convert.ToString(intRandomNumber);
									intRandomNumber = p_utils.RandomNumber(0,9);
									strRandomNumber += "." + Convert.ToString(intRandomNumber);
									intRandomNumber = p_utils.RandomNumber(0,9);
									strRandomNumber += Convert.ToString(intRandomNumber);
									if (strRandomNumber != strRandomSave)
									{
										strRandomSave = strRandomNumber;
										break;
									}
								}
                                strSQL = "insert into travel_time (psite_id,biosum_plot_id,travel_mode,ONE_WAY_HOURS) VALUES ";
								strSQL += "(" + p_ado.m_DataSet.Tables["processing_site"].Rows[y]["psite_id"].ToString()
									+ ",'" + p_ado.m_DataSet.Tables["plot"].Rows[x]["biosum_plot_id"].ToString()
									+ "',1"  +
								       ","  + strRandomNumber + ");";
								p_ado.SqlNonQuery(p_conn,strSQL);
						
							}
							System.Windows.Forms.Application.DoEvents();
							if (this.m_frmTherm.AbortProcess == true) break;
						}
						int intStart=1;
						this.m_frmTherm.progressBar1.Maximum = 
							p_ado.m_DataSet.Tables["processing_site"].Rows.Count;
						strRandomSave="";
						for (x=0;x<=p_ado.m_DataSet.Tables["processing_site"].Rows.Count-1;x++)
						{
							this.m_frmTherm.Increment(x + 1);
							for (y=intStart;y<=p_ado.m_DataSet.Tables["processing_site"].Rows.Count-1;y++)
							{
                                strSQL="";
								if (p_ado.m_DataSet.Tables["processing_site"].Rows[y]["trancd"].ToString().Trim() == "2")
								{
                                    strSQL = "insert into travel_time (psite_id,railhead_id,travel_mode,ONE_WAY_HOURS) values ";
								}
								else if (p_ado.m_DataSet.Tables["processing_site"].Rows[y]["trancd"].ToString().Trim() == "3")
								{
                                    strSQL = "insert into travel_time (psite_id,collector_id,travel_mode,ONE_WAY_HOURS) values ";
								}
								if (strSQL.Trim().Length > 0)
								{
									for (;;)
									{
										intRandomNumber = p_utils.RandomNumber(1,18);
										strRandomNumber = Convert.ToString(intRandomNumber);
										intRandomNumber = p_utils.RandomNumber(0,9);
										strRandomNumber += "." + Convert.ToString(intRandomNumber);
										intRandomNumber = p_utils.RandomNumber(0,9);
										strRandomNumber += Convert.ToString(intRandomNumber);
										if (strRandomNumber != strRandomSave)
										{
											strRandomSave = strRandomNumber;
											break;
										}
									}
									strSQL += "(" + p_ado.m_DataSet.Tables["processing_site"].Rows[x]["psite_id"].ToString()
										+ "," + p_ado.m_DataSet.Tables["processing_site"].Rows[y]["psite_id"].ToString()
										+ ",2" + 
										  ","  + strRandomNumber + ");";
									p_ado.SqlNonQuery(p_conn,strSQL);
									if (p_ado.m_intError != 0)
										break;
								}
								
								
							}
							if (p_ado.m_intError != 0)
								break;
							intStart++;
							System.Windows.Forms.Application.DoEvents();
							if (this.m_frmTherm.AbortProcess == true) break;
						}
						if (p_ado.m_intError != 0)
						{
						}
						else
						{
							MessageBox.Show("Finished Generating Travel Times");
						}
					    this.m_frmTherm.Close();
						this.m_frmTherm = null;
					
					}
					p_ado.m_DataSet.Clear();
					p_ado.m_DataSet = null;
					p_ado.m_OleDbDataAdapter.Dispose();
					p_ado.m_OleDbDataAdapter = null;
					p_conn.Close();
					
					p_ado.m_OleDbConnection=null;

				}
			}
			p_conn=null;
			p_ado=null;
			p_utils=null;
		

     

		}
		private void ThermCancel(object sender, System.EventArgs e)
		{
			string strMsg = "Do you wish to cancel generating travel times (Y/N)?";
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
