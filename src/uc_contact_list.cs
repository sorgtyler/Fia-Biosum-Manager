using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_contact_list.
	/// </summary>
	public class uc_contact_list : System.Windows.Forms.UserControl
	{

		private string m_strProjDir;
		private string m_strFieldList="";
		private FIA_Biosum_Manager.ado_data_access m_ado;
		private string m_strConn;
		private System.Data.DataView m_dv;
		public int m_intIndex=0;
		private int m_intCurrRow=0;
		public int m_intError=0;
		private bool m_bLoaded=false;
		private const int MENU_FILTERBYVALUE=0;
		private const int MENU_FILTERBYENTEREDVALUE=1;
		private const int MENU_REMOVEFILTER = 2;
		private const int MENU_UNIQUEVALUES = 3;
		private const int MENU_MODIFY = 5;
		private const int MENU_DELETE=6;
		private const int MENU_SELECTALL = 8;
		private const int MENU_IDXASC = 10;
		private const int MENU_IDXDESC = 11;
		private const int MENU_REMOVEIDX = 12;
		private const int MENU_MAX = 14;
		private const int MENU_MIN = 15;
		private const int MENU_AVG = 16;
		private const int MENU_SUM = 17;
		private const int MENU_COUNTBYVALUE=18;


		

		private int m_intPopupColumn=0;
		public System.Windows.Forms.ContextMenu m_mnuDataGridPopup;
		private bool m_bDelete=false;
		private string m_strDeletedList="";
		private int m_intDeletedCount=0;
		private string m_strColumnFilterList="";
		private string m_strColumnSortList="";
		private string m_strUserProcesses="";

		private System.Data.DataTable m_dtTableSchema;
		private System.Data.DataRelation m_dataRelation;
		public  ResizeFormUsingVisibleScrollBars m_oResizeForm = new ResizeFormUsingVisibleScrollBars();
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.ComboBox cmbFilter;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnHelp;
		public System.Windows.Forms.Button btnDelete;
		private System.Windows.Forms.Button btnNew;
		private System.Windows.Forms.Button btnEdit;
		public System.Windows.Forms.Button btnSave;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.DataGrid m_dg;
		public System.Windows.Forms.Label lblTitle;

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_contact_list()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			m_oResizeForm.ScrollBarParentControl=panel1;
			
			



			// TODO: Add any initialization after the InitializeComponent call

		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}
		public void loadvalues(string p_strProjDir)
		{
			if (this.m_bLoaded==false)
			{
				string strColumnName="";
				this.m_strProjDir = p_strProjDir;
				this.m_ado = new ado_data_access();
				this.m_strConn = this.m_ado.getMDBConnString(this.m_strProjDir + "\\db\\project.mdb","","");
				this.m_ado.OpenConnection(this.m_strConn);
				if (this.m_ado.m_intError != 0)
				{
					this.m_intError = this.m_ado.m_intError;
					this.m_ado = null;
					return ;

				}
				this.InitializePopup();
				                                         
				this.m_ado.m_DataSet = new DataSet("contacts");
				this.m_ado.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
				this.m_ado.m_strSQL = "select name,organization, work_phone,street_addr,city,state,zip,email from contacts order by name;";
				this.m_dtTableSchema = this.m_ado.getTableSchema(this.m_ado.m_OleDbConnection,
					this.m_ado.m_OleDbTransaction,
					this.m_ado.m_strSQL);

			
			
				if (this.m_ado.m_intError == 0)
				{

					for (int x=0; x<=m_dtTableSchema.Rows.Count-1;x++)
					{
						if (this.m_strFieldList.Trim().Length == 0)
						{
							if (m_dtTableSchema.Rows[x]["columnname"].ToString().Trim().ToUpper() != "NAME")
							{
								this.m_strFieldList = m_dtTableSchema.Rows[x]["columnname"].ToString().Trim();
							}
						}
						else
						{	
							if (m_dtTableSchema.Rows[x]["columnname"].ToString().Trim().ToUpper() != "NAME")
							{
								this.m_strFieldList += "," + m_dtTableSchema.Rows[x]["columnname"].ToString().Trim();
							}
						}
					}
					this.InitializeOleDbTransactionCommands();

					this.m_ado.m_strSQL = "select name,organization,work_phone,street_addr,city,state,zip,email from contacts order by NAME;";
					this.m_ado.m_OleDbCommand = this.m_ado.m_OleDbConnection.CreateCommand();
					this.m_ado.m_OleDbCommand.CommandText = this.m_ado.m_strSQL;
					this.m_ado.m_OleDbDataAdapter.SelectCommand = this.m_ado.m_OleDbCommand;
					this.m_ado.m_OleDbDataAdapter.SelectCommand.Transaction = this.m_ado.m_OleDbTransaction;
					try 
					{

						this.m_ado.m_OleDbDataAdapter.Fill(this.m_ado.m_DataSet,"contacts");
						this.m_dv = new DataView(this.m_ado.m_DataSet.Tables["contacts"]);
				
						this.m_dv.AllowNew = false;       //cannot append new records
						this.m_dv.AllowDelete = false;    //cannot delete records
						this.m_dv.AllowEdit = false;
						this.m_dg.CaptionText = "contacts";

						m_dg.BackgroundColor=frmMain.g_oGridViewBackgroundColor;
						if (frmMain.g_oGridViewFont != null) m_dg.Font=frmMain.g_oGridViewFont;
						m_dg.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
						m_dg.ForeColor = frmMain.g_oGridViewRowForegroundColor;

						this.m_ado.m_DataSet.Tables.Add("user_processes");
						this.m_ado.m_DataSet.Tables["user_processes"].Columns.Add("process",typeof(string));
						this.m_ado.m_DataSet.Tables["user_processes"].Columns.Add("name",typeof(string));

						//define primary keys
						DataColumn[] colPk = new DataColumn[1];
						colPk[0] = this.m_ado.m_DataSet.Tables["contacts"].Columns["name"];
						this.m_ado.m_DataSet.Tables["contacts"].PrimaryKey = colPk;

						colPk[0] = this.m_ado.m_DataSet.Tables["user_processes"].Columns["name"];
						this.m_ado.m_DataSet.Tables["user_processes"].PrimaryKey = colPk;

						
						this.m_dataRelation = new DataRelation("ParentChild",
						    this.m_ado.m_DataSet.Tables["contacts"].Columns["name"],
							this.m_ado.m_DataSet.Tables["user_processes"].Columns["name"]);

						this.m_ado.m_strSQL = "SELECT name,process FROM CONTACTS";
						this.m_ado.SqlQueryReader(this.m_ado.m_OleDbConnection,this.m_ado.m_OleDbTransaction,this.m_ado.m_strSQL);
						if (this.m_ado.m_OleDbDataReader.HasRows)
						{
							while (this.m_ado.m_OleDbDataReader.Read())
							{
								System.Data.DataRow p_row = this.m_ado.m_DataSet.Tables["user_processes"].NewRow();
								p_row["name"] = this.m_ado.m_OleDbDataReader["name"];
								p_row["process"] = this.m_ado.m_OleDbDataReader["process"];
								this.m_ado.m_DataSet.Tables["user_processes"].Rows.Add(p_row);
							}
						}
						this.m_ado.m_OleDbDataReader.Close();


						

						/***********************************************************************************
						 **assign the aColumnTextColumn as type DataGridColoredTextBoxColumn object class
						 ***********************************************************************************/
						Contacts_DataGridColoredTextBoxColumn aColumnTextColumn ;


						/***************************************************************
						 **custom define the grid style
						 ***************************************************************/
						DataGridTableStyle tableStyle = new DataGridTableStyle();

						/***********************************************************************
						 **map the data grid table style to the scenario rx intensity dataset
						 ***********************************************************************/
						tableStyle.MappingName = "contacts";
						tableStyle.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
						tableStyle.ForeColor = frmMain.g_oGridViewRowForegroundColor;
						tableStyle.BackColor = frmMain.g_oGridViewRowBackgroundColor;
						tableStyle.SelectionBackColor = frmMain.g_oGridViewSelectedRowBackgroundColor;

   
						/******************************************************************************
						 **since the dataset has things like field name and number of columns,
						 **we will use those to create new columnstyles for the columns in our grid
						 ******************************************************************************/
						//get the number of columns from the scenario_rx_intensity data set
						int numCols = this.m_ado.m_DataSet.Tables["contacts"].Columns.Count;
                
                    
						/************************************************
						 **loop through all the columns in the dataset	
						 ************************************************/
						for(int i = 0; i < numCols; ++i)
						{
							strColumnName = this.m_ado.m_DataSet.Tables["contacts"].Columns[i].ColumnName;
					
							/******************************************************************
							 **create a new instance of the DataGridColoredTextBoxColumn class
							 ******************************************************************/
							aColumnTextColumn = new Contacts_DataGridColoredTextBoxColumn(false,false,this);
							/***********************************
							 **all columns are read-only except
							 **the edit columns
							 ***********************************/
							aColumnTextColumn.ReadOnly=true;


						

							aColumnTextColumn.HeaderText = strColumnName;

				 				    
							/********************************************************************
							 **assign the mappingname property the data sets column name
							 ********************************************************************/
							aColumnTextColumn.MappingName = strColumnName;
							//aColumnTextColumn
							aColumnTextColumn.TextBox.ContextMenu =  new ContextMenu(); //this.m_mnuDataGridPopup;
							aColumnTextColumn.TextBox.ContextMenu = this.m_mnuDataGridPopup;
							aColumnTextColumn.TextBox.MouseDown += new System.Windows.Forms.MouseEventHandler(this.m_dg_TextBox_MouseDown);
							/********************************************************************
							 **add the datagridcoloredtextboxcolumn object to the data grid 
							 **table style object
							 ********************************************************************/
							tableStyle.GridColumnStyles.Add(aColumnTextColumn);

						}
						/*********************************************************************
						 ** make the dataGrid use our new tablestyle and bind it to our table
						 *********************************************************************/
						this.m_dg.TableStyles.Clear();
						this.m_dg.TableStyles.Add(tableStyle);

						this.m_dg.DataSource = this.m_dv;  

						if (this.m_ado.m_DataSet.Tables["contacts"].Rows.Count > 0)
						{
							this.btnDelete.Enabled=true;
							this.btnEdit.Enabled=true;
						}

				

						this.m_dg.Expand(-1);
					}
					catch (Exception e)
					{
						MessageBox.Show(e.Message,"Table",MessageBoxButtons.OK,MessageBoxIcon.Error);
						this.m_intError=-1;
						this.m_ado.m_OleDbConnection.Close();
						this.m_ado.m_OleDbConnection = null;
						this.m_ado.m_DataSet.Clear();
						this.m_ado.m_DataSet= null;
						this.m_ado.m_OleDbDataAdapter.Dispose();
						this.m_ado.m_OleDbDataAdapter = null;
						this.m_ado = null;
						return;

					}

			
			

					if (this.m_ado.m_DataSet.Tables["contacts"].Rows.Count == 0)
					{
						this.m_intCurrRow = 1;
					}
					else
					{
						this.m_intCurrRow = 1;
					}
			

					//event handler to keep track of current row and cell movement
					this.m_dg.CurrentCellChanged += new
						System.EventHandler(this.m_dg_CurrentCellChanged);

				}
				if (this.m_ado.m_intError < 0)
				{
					this.ParentForm.Close();
				}
				else
				{
					this.m_bLoaded=true;
				}
			}
		}
		private void InitializePopup()
		{
			this.m_mnuDataGridPopup = new ContextMenu();
			this.m_mnuDataGridPopup.MenuItems.Add("Filter By Value",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Filter By Entered Value", new EventHandler(Filter_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Remove Filter",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Unique Values",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=false;
			this.m_mnuDataGridPopup.MenuItems.Add("-");
			this.m_mnuDataGridPopup.MenuItems.Add("Modify",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems[MENU_MODIFY].Enabled=false;
			this.m_mnuDataGridPopup.MenuItems.Add("Delete",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems[MENU_DELETE].Enabled=false;
			this.m_mnuDataGridPopup.MenuItems.Add("-");
			this.m_mnuDataGridPopup.MenuItems.Add("Select All",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems[MENU_SELECTALL].Enabled=true;
			this.m_mnuDataGridPopup.MenuItems.Add("-");
			this.m_mnuDataGridPopup.MenuItems.Add("Index Ascending",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Index Descending",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Remove Index",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=false;
			this.m_mnuDataGridPopup.MenuItems.Add("-");
			this.m_mnuDataGridPopup.MenuItems.Add("Maximum",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Minimum",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Average",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Sum",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Count By Value",new EventHandler(PopUp_Clicked));
			
			

			
		}
		private void PopUp_Clicked(object sender,EventArgs e)
		{
			MenuItem miClicked = (MenuItem)sender;
			string strItem = miClicked.Text;
			string strCellValue = "";
			string strCol = "";
			string strFilter = "";
			string strExp="";
			
			int x=0;
			
			switch (strItem.ToUpper().Trim())
			{
				case "FILTER BY VALUE":

					strCellValue = this.m_dg[this.m_dg.CurrentCell.RowNumber,this.m_dg.CurrentCell.ColumnNumber].ToString();
					this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=true;
					strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
					
				switch (this.m_dg[this.m_dg.CurrentCell.RowNumber,this.m_dg.CurrentCell.ColumnNumber].GetType().FullName.ToString().Trim())
				{
					case "System.String":
						if (strCellValue.Trim().Length==0)
						{
							strFilter = strCol + " IS NULL";
						}
						else

							strFilter =  strCol + "='" + strCellValue + "'";
						break;
					default:
						if (strCellValue.Trim().Length==0)
						{
							strFilter = strCol + " IS NULL";
						}
						else
							strFilter = strCol + "=" + strCellValue;

						break;
				}
					this.m_dv.RowFilter = strFilter;

					if (this.m_strColumnFilterList.Trim().Length > 0) this.m_strColumnFilterList+=  strCol.Trim() + ",";
					else this.m_strColumnFilterList = strCol.Trim() + ",";

					break;
				case "UNIQUE VALUES":
					this.uniquevalues();
					break;
				case "MODIFY":
					break;
				case "REMOVE FILTER":
					this.m_dv.RowFilter="";
					this.cmbFilter.Text = "All";
					this.m_strColumnFilterList="";
					this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=false;
					break;
				case "INDEX ASCENDING":
					this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=true;
					strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
					this.m_dv.Sort = strCol + " ASC";
					this.m_strColumnSortList = strCol.Trim() + ",";
					break;
				case "INDEX DESCENDING":
					this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=true;
					strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
					this.m_dv.Sort = strCol + " DESC";
					this.m_strColumnSortList = strCol.Trim() + ",";
					break;
				case "REMOVE INDEX":
					this.m_dv.Sort="";
					this.m_strColumnSortList="";
					break;
				case "MAXIMUM":
					try
					{
						strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
						strFilter = "Max(" + strCol + ")";
						this.m_dv.RowStateFilter = System.Data.DataViewRowState.CurrentRows;
						if (this.m_dv.RowFilter.ToString().Trim().Length > 0)
						{
							MessageBox.Show(strCol + " Maximum: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString()));
						}
						else
						{
							MessageBox.Show(strCol + " Maximum: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter, null));
						}
					}
					catch
					{
					}
					break;
				case "MINIMUM":
					try
					{
						strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
						strFilter = "Min(" + strCol + ")";
						this.m_dv.RowStateFilter = System.Data.DataViewRowState.CurrentRows;
						if (this.m_dv.RowFilter.ToString().Trim().Length > 0)
						{
							MessageBox.Show(strCol + " Minimum: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString()));
						}
						else
						{
							MessageBox.Show(strCol + " Minimum: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter, null));
						}
					}
					catch
					{
					}
					break;

				case "AVERAGE":
					try
					{
						strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
						strFilter = "Avg(" + strCol + ")";
						this.m_dv.RowStateFilter = System.Data.DataViewRowState.CurrentRows;
						if (this.m_dv.RowFilter.ToString().Trim().Length > 0)
						{
							MessageBox.Show(strCol + " Average: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString()));
						}
						else
						{
							MessageBox.Show(strCol + " Average: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter, null));
						}
					}
					catch
					{
					}
					break;

				case "SUM":
					try
					{
						strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
						strFilter = "Sum(" + strCol + ")";
						this.m_dv.RowStateFilter = System.Data.DataViewRowState.CurrentRows;
						if (this.m_dv.RowFilter.ToString().Trim().Length > 0)
						{
							MessageBox.Show(strCol + " Sum: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString()));
						}
						else
						{
							MessageBox.Show(strCol + " Sum: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter, null));
						}
					}
					catch
					{
					}
					break;

				case "COUNT BY VALUE":
					try
					{
						strCellValue = string.Format("{0:#.####################}",this.m_dg[this.m_dg.CurrentCell.RowNumber,this.m_dg.CurrentCell.ColumnNumber].ToString());
						strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
						switch (this.m_dg[this.m_dg.CurrentCell.RowNumber,this.m_dg.CurrentCell.ColumnNumber].GetType().FullName.ToString().Trim())
						{
							case "System.String":
								strExp =  strCol + "='" + strCellValue + "'";
								break;
							default:
								strExp = strCol + "=" + strCellValue;
								break;
						}
						strFilter = "Count(" + strCol + ")";
						this.m_dv.RowStateFilter = System.Data.DataViewRowState.CurrentRows;
						if (this.m_dv.RowFilter.ToString().Trim().Length > 0)
						{
							MessageBox.Show(strCol + " Count: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString() + " and " + strExp));
						}
						else
						{
							MessageBox.Show(strCol + " Count: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter, strExp));
						}
					}
					catch
					{
					}
					break;
				case "SELECT ALL":
					int intRows = this.m_dg.BindingContext[this.m_dg.DataSource,this.m_dg.DataMember].Count;
					for (x=0; x<=intRows-1;x++)
					{
						this.m_dg.Select(x);
					}
					break;
				case "DELETE":
					this.DeleteRecords();

					break;
					
					
		    
			}
		}
		private void Filter_Clicked(object sender,EventArgs e)
		{
			MenuItem miClicked = (MenuItem)sender;
			string strItem = miClicked.Text;
			string strCellValue = "";
			string strCol = "";
			string strDataType="";
			string strFilter = "";
			
			string strExp="";
			int intLeft=0;
			int intTop=0;
			int x=0;
			int y=0;
			System.Windows.Forms.DialogResult dlgResult;
			System.Windows.Forms.CurrencyManager p_cm;
			System.Data.DataView p_dv;
			int intCurrRow=0;
			//MessageBox.Show(strItem);
			switch (strItem.ToUpper().Trim())
			{
				case "FILTER BY ENTERED VALUE":
					strCellValue = this.m_dg[this.m_dg.CurrentCell.RowNumber,this.m_dg.CurrentCell.ColumnNumber].ToString();
					strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
					strDataType = this.m_dg[this.m_dg.CurrentCell.RowNumber,this.m_dg.CurrentCell.ColumnNumber].GetType().FullName.ToString().Trim();
				switch (strDataType)
				{
					case "System.String":
						frmDialog frmTemp = new frmDialog();
						frmTemp.Initialize_Filter_Rows_Text_Datatype_User_Control();
						frmTemp.Text = strCol;
						frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
						dlgResult = frmTemp.ShowDialog();
						if (dlgResult.ToString().Trim().ToUpper() == "OK" )
						{
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=true;
							string strOperation = frmTemp.uc_filter_rows_text_datatype1.GetFilterType();
							string strText = frmTemp.uc_filter_rows_text_datatype1.GetText();

							switch (strOperation)
							{
								case "EQUAL":
									if (strText.Trim().Length==0)
										strFilter = strCol + " IS NULL";
									else
										strFilter =  strCol + "='" + strText + "'";
									break;
								case "NOTEQUAL":
									if (strText.Trim().Length==0)
										strFilter = strCol + " IS NOT NULL";
									else
										strFilter =  strCol + "<>'" + strText + "'";
									break;
								case "START":
									if (strText.Trim().Length==0)
										strFilter = strCol + " IS NULL";
									else
										strFilter =  strCol + " LIKE '" + strText.Trim() + "*'";
									break;
								case "NOTSTART":
									if (strText.Trim().Length==0)
										strFilter = strCol + " IS NOT NULL";
									else
										strFilter =  strCol + " NOT LIKE '" + strText.Trim() + "*'";
									break;
								case "CONTAIN":
									if (strText.Trim().Length==0)
										strFilter = strCol + " IS NULL";
									else
										strFilter =  strCol + " LIKE '%" + strText.Trim() + "%'";
									break;
								case "NOTCONTAIN":
									if (strText.Trim().Length==0)
										strFilter = strCol + " IS NOT NULL";
									else
										strFilter =  strCol + " NOT LIKE '%" + strText.Trim() + "%'";
									break;
								case "END":
									if (strText.Trim().Length==0)
										strFilter = strCol + " IS NULL";
									else
										strFilter =  strCol + " LIKE '*" + strText.Trim() + "'";
									break;
								case "NOTEND":
									if (strText.Trim().Length==0)
										strFilter = strCol + " IS NULL";
									else
										strFilter =  strCol + " NOT LIKE '*" + strText.Trim() + "'";
									break;                                    
							}
						}
						break;
					default:
						if (strDataType=="System.Integer" || strDataType=="System.Double" ||
							strDataType=="System.Int16" || strDataType=="System.Int32" ||
							strDataType=="System.Int64" || strDataType=="System.Byte" ||
							strDataType=="System.Decimal")
						{
							frmDialog frmTemp2 = new frmDialog();
							frmTemp2.Initialize_Filter_Rows_Numeric_Datatype_User_Control();

							frmTemp2.Text = strCol;
							frmTemp2.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
							dlgResult = frmTemp2.ShowDialog();
							if (dlgResult.ToString().Trim().ToUpper() == "OK" )
							{
								this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=true;
								string strOperation = frmTemp2.uc_filter_rows_numeric_datatype1.GetFilterType();
								string strText = frmTemp2.uc_filter_rows_numeric_datatype1.GetSmallestText().Replace(",","");
								string strLargestText = frmTemp2.uc_filter_rows_numeric_datatype1.GetLargestText().Replace(",","");

								switch (strOperation)
								{
									case "EQUAL":
										if (strText.Trim().Length==0)
											strFilter = strCol + " IS NULL";
										else
											strFilter =  strCol + "=" + strText;
										break;
									case "NOTEQUAL":
										if (strText.Trim().Length==0)
											strFilter = strCol + " IS NOT NULL";
										else
											strFilter =  strCol + "<>" + strText;
										break;
									case "GREATERTHAN":
										if (strText.Trim().Length==0)
											strFilter = strCol + " IS NULL";
										else
											strFilter =  strCol + " > " + strText.Trim();
										break;
									case "LESSTHAN":
										if (strText.Trim().Length==0)
											strFilter = strCol + " IS NULL";
										else
											strFilter =  strCol + " < " + strText.Trim();
										break;
									case "BETWEEN":
										if (strText.Trim().Length==0)
											strFilter = strCol + " IS NULL";
										else
											strFilter =  strCol + " >= " + strText + " AND " + strCol + " <= " + strLargestText;
										break;
										              
								}
							}
						}
						break;

				}
			
					break;
			}
			if (strFilter.Trim().Length > 0)
			{
				this.m_dv.RowFilter = strFilter;
				
				if (this.m_strColumnFilterList.Trim().Length > 0) this.m_strColumnFilterList+=  strCol.Trim() + ",";
				else this.m_strColumnFilterList = strCol.Trim() + ",";
			}
			
					
			
		}

		public void DeleteRecords()
		{
			int x;
			int intCurrRow;
			string strValue="";
			System.Windows.Forms.CurrencyManager p_cm;
			System.Data.DataView p_dv;
			string strList="";

			this.m_intError=0;

			p_cm = (CurrencyManager)this.BindingContext[this.m_dg.DataSource,this.m_dg.DataMember];
			p_dv = (DataView)p_cm.List;
			intCurrRow= this.m_intCurrRow-1;
			//add selected plots to the delete list
			for (x=0; x < p_dv.Count;x++)
			{
				if (m_dg.IsSelected(x) || x == intCurrRow)
				{
					if (this.m_strDeletedList.Trim().Length > 0)
					{
						this.m_strDeletedList += "," +  p_dv[x]["name"].ToString().Trim().ToUpper();
					}
					else
					{
						this.m_strDeletedList = p_dv[x]["name"].ToString().Trim().ToUpper();
					}
					strList += "," + p_dv[x]["name"].ToString().Trim().ToUpper() + ",";
					this.m_intDeletedCount++;
						
				}
			}
			if (this.m_intDeletedCount > 0)
			{
				string strMsg = this.m_intDeletedCount.ToString().Trim() + " record(s) will be deleted from the table. \n Do you want to delete the record(s) from the table? (Y/N)";
				DialogResult result = MessageBox.Show(strMsg,"Delete Selected Processing Sites", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
				if (result == DialogResult.Yes)
				{
					this.m_dv.AllowDelete = true;

					if (this.m_intError==0 && this.m_ado.m_intError==0)
					{
						//delete any plots that are in the delete list
						for (x=0; x < p_dv.Count;x++)
						{
							strValue = "," + p_dv[x]["name"].ToString().Trim().ToUpper() + ",";
							if (strList.IndexOf(strValue) >=0)
							{
								p_dv[x].Delete();
								x=x-1;
							}
						}
						if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
					}
				}
			}
			this.m_strDeletedList="";
			this.m_intDeletedCount=0;
			this.m_dv.AllowDelete=false;
			
			return;
		}


		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.panel1 = new System.Windows.Forms.Panel();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbFilter = new System.Windows.Forms.ComboBox();
			this.btnClose = new System.Windows.Forms.Button();
			this.btnHelp = new System.Windows.Forms.Button();
			this.btnDelete = new System.Windows.Forms.Button();
			this.btnNew = new System.Windows.Forms.Button();
			this.btnEdit = new System.Windows.Forms.Button();
			this.btnSave = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.m_dg = new System.Windows.Forms.DataGrid();
			this.lblTitle = new System.Windows.Forms.Label();
			this.panel1.SuspendLayout();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.m_dg)).BeginInit();
			this.SuspendLayout();
			// 
			// panel1
			// 
			this.panel1.AutoScroll = true;
			this.panel1.Controls.Add(this.groupBox1);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel1.Location = new System.Drawing.Point(0, 0);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(704, 352);
			this.panel1.TabIndex = 0;
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.cmbFilter);
			this.groupBox1.Controls.Add(this.btnClose);
			this.groupBox1.Controls.Add(this.btnHelp);
			this.groupBox1.Controls.Add(this.btnDelete);
			this.groupBox1.Controls.Add(this.btnNew);
			this.groupBox1.Controls.Add(this.btnEdit);
			this.groupBox1.Controls.Add(this.btnSave);
			this.groupBox1.Controls.Add(this.btnCancel);
			this.groupBox1.Controls.Add(this.m_dg);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(704, 352);
			this.groupBox1.TabIndex = 1;
			this.groupBox1.TabStop = false;
			// 
			// cmbFilter
			// 
			this.cmbFilter.Items.AddRange(new object[] {
														   "All",
														   "Core Analysis",
														   "Frcs",
														   "Fvs",
														   "Gis",
														   "Processor"});
			this.cmbFilter.Location = new System.Drawing.Point(8, 280);
			this.cmbFilter.Name = "cmbFilter";
			this.cmbFilter.Size = new System.Drawing.Size(104, 21);
			this.cmbFilter.TabIndex = 59;
			this.cmbFilter.Text = "All";
			this.cmbFilter.Click += new System.EventHandler(this.cmbFilter_SelectedIndexChanged);
			this.cmbFilter.SelectedIndexChanged += new System.EventHandler(this.cmbFilter_SelectedIndexChanged);
			// 
			// btnClose
			// 
			this.btnClose.Location = new System.Drawing.Point(598, 312);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(96, 32);
			this.btnClose.TabIndex = 58;
			this.btnClose.Text = "Close";
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// btnHelp
			// 
			this.btnHelp.Location = new System.Drawing.Point(12, 312);
			this.btnHelp.Name = "btnHelp";
			this.btnHelp.Size = new System.Drawing.Size(96, 32);
			this.btnHelp.TabIndex = 57;
			this.btnHelp.Text = "Help";
			// 
			// btnDelete
			// 
			this.btnDelete.Enabled = false;
			this.btnDelete.Location = new System.Drawing.Point(336, 280);
			this.btnDelete.Name = "btnDelete";
			this.btnDelete.Size = new System.Drawing.Size(64, 32);
			this.btnDelete.TabIndex = 56;
			this.btnDelete.Text = "Delete";
			this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
			// 
			// btnNew
			// 
			this.btnNew.Location = new System.Drawing.Point(208, 280);
			this.btnNew.Name = "btnNew";
			this.btnNew.Size = new System.Drawing.Size(64, 32);
			this.btnNew.TabIndex = 52;
			this.btnNew.Text = "New";
			this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
			// 
			// btnEdit
			// 
			this.btnEdit.Location = new System.Drawing.Point(272, 280);
			this.btnEdit.Name = "btnEdit";
			this.btnEdit.Size = new System.Drawing.Size(64, 32);
			this.btnEdit.TabIndex = 53;
			this.btnEdit.Text = "Edit";
			this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
			// 
			// btnSave
			// 
			this.btnSave.Enabled = false;
			this.btnSave.Location = new System.Drawing.Point(400, 280);
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(64, 32);
			this.btnSave.TabIndex = 54;
			this.btnSave.Text = "Save";
			this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
			// 
			// btnCancel
			// 
			this.btnCancel.Location = new System.Drawing.Point(464, 280);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(64, 32);
			this.btnCancel.TabIndex = 55;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// m_dg
			// 
			this.m_dg.DataMember = "";
			this.m_dg.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.m_dg.Location = new System.Drawing.Point(8, 64);
			this.m_dg.Name = "m_dg";
			this.m_dg.Size = new System.Drawing.Size(688, 208);
			this.m_dg.TabIndex = 27;
			this.m_dg.CurrentCellChanged += new System.EventHandler(this.m_dg_CurrentCellChanged);
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(698, 32);
			this.lblTitle.TabIndex = 26;
			this.lblTitle.Text = "Contacts";
			// 
			// uc_contact_list
			// 
			this.Controls.Add(this.panel1);
			this.Name = "uc_contact_list";
			this.Size = new System.Drawing.Size(704, 352);
			this.Resize += new System.EventHandler(this.uc_contact_list_Resize);
			this.panel1.ResumeLayout(false);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.m_dg)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void m_dg_CurrentCellChanged(object sender, System.EventArgs e)
		{
			/*********************************************************************
			 **this one command causes the whole row to be highlighted with 
			 **the selected row background and foreground color
			 *********************************************************************/
			if (m_dg.CurrentRowIndex >=0) m_dg.Select(m_dg.CurrentRowIndex);

			if (this.m_intCurrRow > 0)
			{
				if (this.m_dg.CurrentRowIndex != this.m_intCurrRow - 1)
				{
					this.m_intCurrRow = this.m_dg.CurrentRowIndex + 1;
				}
			}
		}

		private void m_dg_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			switch (e.Button)
			{
				case System.Windows.Forms.MouseButtons.Right :
					if (this.m_dv.RowFilter.Trim().Length > 0)
					{
						if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=true;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled==true)
							this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=false;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled==true)
							this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=false;

					}
					else
					{
						if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=true;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=true;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled==true)
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=false;

					}
					if (this.m_dv.Sort.Trim().Length > 0)
					{
						if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=true;
						
					}
					else
					{
						if (this.m_mnuDataGridPopup.MenuItems[MENU_IDXASC].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_IDXASC].Enabled=true;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_IDXDESC].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_IDXDESC].Enabled=true;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled==true)
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=false;

					}
					
					if (this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled==false)
						this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled=true;
					if (this.m_mnuDataGridPopup.MenuItems[MENU_SELECTALL].Enabled==false)
						this.m_mnuDataGridPopup.MenuItems[MENU_SELECTALL].Enabled=true;
					if (this.m_mnuDataGridPopup.MenuItems[MENU_MIN].Enabled==false)
						this.m_mnuDataGridPopup.MenuItems[MENU_MIN].Enabled=true;
					if (this.m_mnuDataGridPopup.MenuItems[MENU_MAX].Enabled==false)
						this.m_mnuDataGridPopup.MenuItems[MENU_MAX].Enabled=true;
					if (this.m_bDelete==true)
					{
						if (this.m_mnuDataGridPopup.MenuItems[MENU_DELETE].Enabled==false)
						{
						}
					}
					   
					break;
					
			}

		}
		private void uniquevalues()
		{
			string strCellValue;
			string strBuild="";
			int count=0;
			int x=0;
			
			string strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();

			System.Windows.Forms.CurrencyManager p_cm = (CurrencyManager)this.BindingContext[this.m_dg.DataSource,this.m_dg.DataMember];
			System.Data.DataView p_dv = (DataView)p_cm.List;
			string[] strValues = new string[p_dv.Count];

			for (x = 0; x<= p_dv.Count-1; x++)
			{
				strCellValue = this.m_dg[x,this.m_dg.CurrentCell.ColumnNumber].ToString();
				if (strBuild.IndexOf(strCellValue,0,strBuild.Length) == -1)
				{
					strValues[count] = strCellValue;
					
					strBuild = strBuild + "'" + strCellValue + "',";
					count++;
				}

			}																																					 
 					
			
			FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
			System.Windows.Forms.ListBox p_lst = new ListBox();
         
						
			//define form properties
			frmTemp.Width = 200;
			frmTemp.Height = 300;
			frmTemp.WindowState = System.Windows.Forms.FormWindowState.Normal;
			frmTemp.MaximizeBox = false;
			frmTemp.MinimizeBox = false;
			frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmTemp.Text = "Unique Values (" + strCol + ")";
		                    
			//define numeric text class properties
			p_lst.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			p_lst.Name = "lstUniqueValues";
			p_lst.TabIndex = 0;
			p_lst.Tag = "";
			p_lst.Enabled = true;
			frmTemp.Controls.Add(p_lst);  //add the text box to the form
			

			p_lst.Top = 5;
			p_lst.Width = frmTemp.ClientSize.Width - 20;
			p_lst.Left = 10;
			p_lst.Items.Clear();
			int intLeft = p_lst.Left;
			int intTop = p_lst.Top;

					
			//----------------OK button-----------------//
			System.Windows.Forms.Button btnTempOK = new Button();
			btnTempOK.Location = new System.Drawing.Point(392, 328);
			btnTempOK.Name = "btnOK";
			btnTempOK.Size = new System.Drawing.Size(80, 32);
			btnTempOK.TabIndex = 2;
			btnTempOK.Text = "OK";
			btnTempOK.Click += new System.EventHandler(frmTemp.btnOK_Click);
			frmTemp.Controls.Add(btnTempOK);
			btnTempOK.Left = (int)(frmTemp.Width * .50)  - (int)(btnTempOK.Width * .50);
			p_lst.Height = frmTemp.ClientSize.Height - btnTempOK.Height -15;
			btnTempOK.Top = intTop + p_lst.Height; // + 10;
			p_lst.Visible = true;
			btnTempOK.Visible=true;
			

			frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			for (x = 0; x<= count - 1; x++)
				p_lst.Items.Add(strValues[x].ToString().Trim());

			frmTemp.ShowDialog();
			strValues = null;
			frmTemp = null;
		}
		/// <summary>
		/// return the row of the containing the column definitions of 
		/// column name, datatype, length, default value, allow nulls, etc.
		/// </summary>
		/// <param name="strColumnName">column definition to look up</param>
		/// <returns></returns>
		private int getTableSchemaColumnDefinition(string strColumnName)
		{
			int x;
			for (x=0;x<=this.m_dtTableSchema.Rows.Count-1;x++)
			{
				if (this.m_dtTableSchema.Rows[x]["COLUMNNAME"].ToString().Trim().ToUpper()==
					strColumnName.Trim().ToUpper())
				{
					return x;
					
				}

			}
			//could not find the column
			return -1;
		}

		private void m_dg_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			switch (e.Button)
			{
				case System.Windows.Forms.MouseButtons.Right :
					Point pt = new Point(e.X,e.Y);
					try
					{
					
                     
						System.Windows.Forms.DataGrid.HitTestInfo hitTestInfo = this.m_dg.HitTest(pt);
						this.m_intPopupColumn = hitTestInfo.Column;
						switch (this.m_dg[0,this.m_intPopupColumn].GetType().FullName.ToString().Trim())
						{
							case "System.String":
								this.m_mnuDataGridPopup.MenuItems[MENU_AVG].Enabled=false;
								this.m_mnuDataGridPopup.MenuItems[MENU_SUM].Enabled=false;
								break;
							default:
								this.m_mnuDataGridPopup.MenuItems[MENU_AVG].Enabled=true;
								this.m_mnuDataGridPopup.MenuItems[MENU_SUM].Enabled=true;
								break;
						}

						switch (hitTestInfo.Type)
						{
							case DataGrid.HitTestType.ColumnHeader:
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=false;
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=false;
								this.m_mnuDataGridPopup.MenuItems[MENU_MODIFY].Enabled=false;
								this.m_mnuDataGridPopup.MenuItems[MENU_COUNTBYVALUE].Enabled=false;

								this.m_mnuDataGridPopup.Show(this, new Point(pt.X,pt.Y));
								break;
							case DataGrid.HitTestType.Cell:
							
								this.m_dg.CurrentCell = new DataGridCell(hitTestInfo.Row,hitTestInfo.Column);
								if (this.m_dg.TableStyles[0].GridColumnStyles[this.m_dg.CurrentCell.ColumnNumber].ReadOnly==true)
								{
									this.m_mnuDataGridPopup.MenuItems[MENU_MODIFY].Enabled=false;
								}
								else
								{
									this.m_mnuDataGridPopup.MenuItems[MENU_MODIFY].Enabled=true;
								}

								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=true;
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=true;
								this.m_mnuDataGridPopup.MenuItems[MENU_COUNTBYVALUE].Enabled=true;
								this.m_mnuDataGridPopup.Show(this, new Point(pt.X,pt.Y + this.m_dg.Top));
								break;
					  
						}

						if (this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled=true;


						if (this.m_dv.RowFilter.Trim().Length > 0)
						{
							if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled==false)
								this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=true;
							if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled==true)
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=false;
							if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled==true)
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=false;

						}
						else
						{
							if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled==false)
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=true;
							if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled==false)
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=true;
							if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled==true)
								this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=false;

						}
						if (this.m_dv.Sort.Trim().Length > 0)
						{
							if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled==false)
								this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=true;
						
						}
						else
						{
							if (this.m_mnuDataGridPopup.MenuItems[MENU_IDXASC].Enabled==false)
								this.m_mnuDataGridPopup.MenuItems[MENU_IDXASC].Enabled=true;
							if (this.m_mnuDataGridPopup.MenuItems[MENU_IDXDESC].Enabled==false)
								this.m_mnuDataGridPopup.MenuItems[MENU_IDXDESC].Enabled=true;
							if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled==true)
								this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=false;

						}
                        
					
					}
					catch
					{
						this.m_mnuDataGridPopup.MenuItems[MENU_COUNTBYVALUE].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_AVG].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_SUM].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_MAX].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_MIN].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_MODIFY].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_SELECTALL].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_IDXASC].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_IDXDESC].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_DELETE].Enabled=false;

						this.m_mnuDataGridPopup.Show(this, new Point(pt.X,pt.Y + this.m_dg.Top));
					}

					break;
				case System.Windows.Forms.MouseButtons.Left :
					Point pt2 = new Point(e.X,e.Y);
					try
					{
					
						// DataGrid.HitTestInfo hitTestInfo ;
                     
						System.Windows.Forms.DataGrid.HitTestInfo hitTestInfo = this.m_dg.HitTest(pt2);
						this.m_intPopupColumn = hitTestInfo.Column;

						switch (hitTestInfo.Type)
						{
							case DataGrid.HitTestType.ColumnHeader:
								this.m_strColumnSortList = 
									this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim() + ",";
								break;
							case DataGrid.HitTestType.Cell:
								break;
						}

					
					}
					catch
					{
					}

					break;

			}
		}
		/// <summary>
		/// configure the data adapter to insert, update, and delete data using oledb sql
		/// </summary>
		private void InitializeOleDbTransactionCommands()
		{
			this.m_ado.m_strSQL = "select name,organization,street_addr,city,zip,email,work_phone,state from contacts order by name;";
			//initialize the transaction object with the connection
			this.m_ado.m_OleDbTransaction = this.m_ado.m_OleDbConnection.BeginTransaction();

			this.m_ado.ConfigureDataAdapterInsertCommand(this.m_ado.m_OleDbConnection,
				this.m_ado.m_OleDbDataAdapter,
				this.m_ado.m_OleDbTransaction,
				this.m_ado.m_strSQL,
				"contacts");

			this.m_ado.m_strSQL = "select " + this.m_strFieldList  +  " from contacts order by name;";
			this.m_ado.ConfigureDataAdapterUpdateCommand(this.m_ado.m_OleDbConnection,
				this.m_ado.m_OleDbDataAdapter,
				this.m_ado.m_OleDbTransaction,
				this.m_ado.m_strSQL,"select name from contacts",
				"contacts");

			this.m_ado.m_strSQL = "select " + this.m_strFieldList  +  " from contacts order by name;";
			this.m_ado.ConfigureDataAdapterDeleteCommand(this.m_ado.m_OleDbConnection,
				this.m_ado.m_OleDbDataAdapter,
				this.m_ado.m_OleDbTransaction,
				"select name from contacts",
				"contacts");

			


		}
		private void m_dg_TextBox_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			switch (e.Button)
			{
				case System.Windows.Forms.MouseButtons.Right :

					
					this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=true;
					this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=true;
					this.m_mnuDataGridPopup.MenuItems[MENU_COUNTBYVALUE].Enabled=true;
					Point pt = new Point(e.X,e.Y);
					this.m_intPopupColumn=this.m_dg.CurrentCell.ColumnNumber;
					if (this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].ReadOnly==true)
					{
						this.m_mnuDataGridPopup.MenuItems[MENU_MODIFY].Enabled=false;
					}
				switch (this.m_dg[0,this.m_intPopupColumn].GetType().FullName.ToString().Trim())
				{
					case "System.String":
						this.m_mnuDataGridPopup.MenuItems[MENU_AVG].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_SUM].Enabled=false;

						break;
					default:
						this.m_mnuDataGridPopup.MenuItems[MENU_AVG].Enabled=true;
						this.m_mnuDataGridPopup.MenuItems[MENU_SUM].Enabled=true;
						break;
				}
					if (this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled==false)
						this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled=true;

					if (this.m_dv.RowFilter.Trim().Length > 0)
					{
						if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=true;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled==true)
							this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=false;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled==true)
							this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=false;

					}
					else
					{
						if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=true;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=true;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled==true)
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=false;

					}
					if (this.m_dv.Sort.Trim().Length > 0)
					{
						if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=true;
						
					}
					else
					{
						if (this.m_mnuDataGridPopup.MenuItems[MENU_IDXASC].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_IDXASC].Enabled=true;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_IDXDESC].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_IDXDESC].Enabled=true;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled==true)
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=false;

					}



					break;
					
			}

			
																	   
		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			
			this.ParentForm.Close();
		}

		private void btnNew_Click(object sender, System.EventArgs e)
		{
			this.EditForm("NEW");

		}
		/// <summary>
		/// validate the input objects located on the p_groupBox object
		/// </summary>
		/// <param name="p_groupBox">groupbox control that owns the input objects</param>
		private void val_data(System.Windows.Forms.GroupBox p_groupBox)
		{
		}

		private void btnEdit_Click(object sender, System.EventArgs e)
		{
			this.EditForm("Edit");

		}

		/// <summary>
		/// find the table field in the grid by matching the field name with the column header
		/// </summary>
		/// <param name="strColumn">Column to search for</param>
		/// <returns>grid column number</returns>
		private int getGridColumn(string strColumn)
		{
			int x;
			int intGridCol=-1;
			for (x=0;x<=this.m_dg.TableStyles[0].GridColumnStyles.Count-1;x++)
			{
				if (strColumn.Trim().ToUpper() == this.m_dg.TableStyles[0].GridColumnStyles[x].HeaderText.Trim().ToUpper())
				{
					intGridCol=x;
					break;
				}
			}
			return intGridCol;
		}

		private void uc_contact_list_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
				this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
				this.btnHelp.Top = this.btnClose.Top;
				this.m_dg.Width = this.groupBox1.Width  - (this.m_dg.Left * 2);
				this.m_dg.Height = this.groupBox1.Height - this.m_dg.Top - this.btnClose.Height - this.btnNew.Height - 50;
				this.btnNew.Top =this.m_dg.Height + this.m_dg.Top + 5;
				this.btnEdit.Top = this.btnNew.Top;
				this.btnSave.Top = this.btnNew.Top;
				this.btnDelete.Top = this.btnNew.Top;
				this.btnCancel.Top = this.btnNew.Top;
				this.btnDelete.Left = (int)(this.m_dg.Width * .50) - (int)(this.btnDelete.Width * .50);
				this.btnEdit.Left = this.btnDelete.Left - this.btnEdit.Width;
				this.btnNew.Left = this.btnEdit.Left - this.btnNew.Width;
				this.btnSave.Left = this.btnDelete.Left + this.btnSave.Width;
				this.btnCancel.Left = this.btnSave.Left + this.btnCancel.Width;
				this.cmbFilter.Left = this.m_dg.Left;
				this.cmbFilter.Top = this.btnNew.Top;

			}
			catch
			{
			}
		}

		private void btnSave_Click(object sender, System.EventArgs e)
		{
			this.savevalues();
		}
		public void savevalues()
		{
			int intCurrRow;
			bool bInitializeTransact=false;
		
			

			this.m_intError=0;

			/******************************************************
			 **save the current row, move the current row to a
			 **different row to enable getchanges() method, then
			 **move back to current row
			 ******************************************************/
			intCurrRow = this.m_dg.CurrentRowIndex;
			if (intCurrRow==0)
			{
				this.m_dg.CurrentRowIndex++;
			}
			else
			{
				this.m_dg.CurrentRowIndex=0;
			}
			 
               			
			System.Data.DataTable p_dtChanges;
			System.Data.DataRow p_rowFound;

			try
			{
				
				p_dtChanges = this.m_ado.m_DataSet.Tables["contacts"].GetChanges();
								
				//check if any inserted rows
				
				if (p_dtChanges.HasErrors)
				{
					this.m_ado.m_DataSet.Tables["contacts"].RejectChanges();
					this.m_intError=-1;
				}
				else
				{
					this.m_ado.m_OleDbDataAdapter.Update(this.m_ado.m_DataSet.Tables["contacts"]);
					this.m_ado.m_OleDbTransaction.Commit();
					this.m_ado.m_DataSet.Tables["contacts"].AcceptChanges();
					bInitializeTransact = true;
				}
				if (this.m_intError==0)
				{

					p_dtChanges = this.m_ado.m_DataSet.Tables["user_processes"].GetChanges();
								
					if (p_dtChanges.HasErrors)
					{
						this.m_ado.m_DataSet.Tables["contacts"].RejectChanges();
						this.m_intError=-1;
					}
					else
					{
						//loop through all the rows that changed
						for (int x=0;x<=p_dtChanges.Rows.Count-1;x++)
						{
							//see if the parent row exists
							System.Object[] p_search = new Object[1];
							p_search[0] = p_dtChanges.Rows[x]["name"].ToString();
							p_rowFound = this.m_ado.m_DataSet.Tables["contacts"].Rows.Find(p_search);
							if (p_rowFound != null)
							{
								this.m_ado.m_strSQL = "UPDATE contacts SET process = '" + 
									                      p_dtChanges.Rows[x]["process"].ToString().Trim() + "' " + 
									                  "WHERE trim(ucase(name)) = '" + p_dtChanges.Rows[x]["name"].ToString().Trim().ToUpper() + "'";
								this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);

							}
						}
					}
				}
				if (this.m_intError==0 && bInitializeTransact == true)
				{
					this.InitializeOleDbTransactionCommands();
				}
					
					
				

				
			}
			catch (Exception caught)
			{
				this.m_intError=-1;
				MessageBox.Show(caught.Message);
				this.m_ado.m_DataSet.Tables["contacts"].RejectChanges();
				//rollback the transaction to the original records 
				this.m_ado.m_OleDbTransaction.Rollback();
				
			}
			

			p_dtChanges=null;
			


			this.m_dg.CurrentRowIndex = intCurrRow;		
			this.btnSave.Enabled=false;	
            
			return;
		}

		private void btnDelete_Click(object sender, System.EventArgs e)
		{
			this.DeleteRecords();
		}

		private void EditForm(string p_strAction)
		{
			//check to see if edit and no records to edit
			if (p_strAction.Trim().ToUpper() != "NEW" && this.m_dg.BindingContext[this.m_dg.DataSource,this.m_dg.DataMember].Count==0)
			{
				return;
			}
			System.Data.DataRow  p_rowFound;
			FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
			frmTemp.MaximizeBox = false;
			frmTemp.BackColor = System.Drawing.SystemColors.Control;
			frmTemp.Text = "Project: Contacts (" + p_strAction + ")";

			FIA_Biosum_Manager.uc_contact_edit  p_uc = new uc_contact_edit(this.m_ado,"contacts");
			
			frmTemp.Controls.Add(p_uc);
			frmTemp.ContactsEditUserControl=p_uc ;

			frmTemp.Height=0;
			frmTemp.Width=0;
			if (p_uc.Top + p_uc.Height > frmTemp.ClientSize.Height + 2)
			{
				for (int x=1;;x++)
				{
					frmTemp.Height = x;
					if (p_uc.Top + 
						p_uc.Height < 
						frmTemp.ClientSize.Height)
					{
						break;
					}
				}

			}
			if (p_uc.Left + p_uc.Width > frmTemp.ClientSize.Width + 2)
			{
				for (int x=1;;x++)
				{
					frmTemp.Width = x;
					if (p_uc.Left + 
						p_uc.Width < 
						frmTemp.ClientSize.Width)
					{
						break;
					}
				}

			}
			frmTemp.DisposeOfFormWhenClosing = false;
			p_uc.Dock = System.Windows.Forms.DockStyle.Fill;
			frmTemp.Left = 0;
			frmTemp.Top = 0;
			frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			try
			{
				if (p_strAction.Trim().ToUpper() == "NEW")
				{
					//p_uc.strId = Convert.ToString(this.getUniqueId());
				}
				else
				{
					p_uc.strCity = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("city")].ToString().Trim();
					p_uc.strState = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("state")].ToString().Trim();
					p_uc.strName = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("name")].ToString().Trim();
					p_uc.strEmail = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("email")].ToString().Trim();
					p_uc.strZipCode = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("zip")].ToString().Trim();
					p_uc.strPhoneNumber = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("work_phone")].ToString().Trim();
					p_uc.strStreetAddress = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("street_addr")].ToString().Trim();
					p_uc.strOrganization = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("organization")].ToString().Trim();

					System.Object[] p_search = new Object[1];
					p_search[0] = p_uc.strName.Trim();
					p_rowFound = this.m_ado.m_DataSet.Tables["user_processes"].Rows.Find(p_search);
					if (p_rowFound != null)
					{
						string str = p_rowFound["process"].ToString().Trim().ToUpper();
						if (str.IndexOf("&CORE&") >=0)
						{
							p_uc.checkCore = true;
						}
						if (str.IndexOf("&FVS&") >=0)
						{
							p_uc.checkFvs = true;
						}
						if (str.IndexOf("&FRCS&") >=0)
						{
							p_uc.checkFrcs = true;
						}
						if (str.IndexOf("&PROCESSOR&") >=0)
						{
							p_uc.checkProcessor = true;
						}
						if (str.IndexOf("&GIS&") >=0)
						{
							p_uc.checkGis = true;
						}

					}
					

					

				}
				System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
				if (result==System.Windows.Forms.DialogResult.OK)
				{
					if (p_strAction.Trim().ToUpper()=="NEW")
					{
						this.m_dv.AllowNew = true;
						System.Data.DataRow p_row =	this.m_ado.m_DataSet.Tables["contacts"].NewRow();
						p_row["name"] = p_uc.strName;
						p_row["street_addr"] = p_uc.strStreetAddress;
						p_row["work_phone"] = p_uc.strPhoneNumber;
						p_row["city"] = p_uc.strCity;
						p_row["zip"] = p_uc.strZipCode;
						p_row["state"] = p_uc.strState;
						p_row["email"] = p_uc.strEmail;
						p_row["organization"] = p_uc.strOrganization;
						this.m_ado.m_DataSet.Tables["contacts"].Rows.Add(p_row);
						p_row=null;
						this.m_dv.AllowNew = false;
						this.m_strUserProcesses="";
						if (p_uc.checkCore==true) this.m_strUserProcesses = "&CORE&";
						if (p_uc.checkFrcs==true) this.m_strUserProcesses += "&FRCS&";
						if (p_uc.checkFvs==true) this.m_strUserProcesses += "&FVS&";
						if (p_uc.checkProcessor==true) this.m_strUserProcesses += "&PROCESSOR&";
                        if (p_uc.checkGis==true) this.m_strUserProcesses += "&GIS&";

						p_row =	this.m_ado.m_DataSet.Tables["user_processes"].NewRow();
						p_row["name"] = p_uc.strName;
						p_row["process"] = this.m_strUserProcesses;
						this.m_ado.m_DataSet.Tables["user_processes"].Rows.Add(p_row);
					}
					else
					{
						
						this.m_dg[this.m_intCurrRow-1,this.getGridColumn("name")] = p_uc.strName;
						this.m_dg[this.m_intCurrRow-1,this.getGridColumn("email")] = p_uc.strEmail;
						this.m_dg[this.m_intCurrRow-1,this.getGridColumn("street_addr")] = p_uc.strStreetAddress;
						this.m_dg[this.m_intCurrRow-1,this.getGridColumn("city")] = p_uc.strCity;
						this.m_dg[this.m_intCurrRow-1,this.getGridColumn("state")] = p_uc.strState;
						this.m_dg[this.m_intCurrRow-1,this.getGridColumn("zip")] = p_uc.strZipCode;
						this.m_dg[this.m_intCurrRow-1,this.getGridColumn("work_phone")] = p_uc.strPhoneNumber;
						this.m_dg[this.m_intCurrRow-1,this.getGridColumn("organization")] = p_uc.strOrganization;
						this.m_dg.SetDataBinding(this.m_dv,"");
						this.m_dg.Update();

						System.Object[] p_search = new Object[1];
						p_search[0] = p_uc.strName.Trim();
						p_rowFound = this.m_ado.m_DataSet.Tables["user_processes"].Rows.Find(p_search);
						if (p_rowFound != null)
						{
							this.m_strUserProcesses="";
							if (p_uc.checkCore==true) this.m_strUserProcesses = "&CORE&";
							if (p_uc.checkFrcs==true) this.m_strUserProcesses += "&FRCS&";
							if (p_uc.checkFvs==true) this.m_strUserProcesses += "&FVS&";
							if (p_uc.checkProcessor==true) this.m_strUserProcesses += "&PROCESSOR&";
							if (p_uc.checkGis==true) this.m_strUserProcesses += "&GIS&";
							p_rowFound.BeginEdit();
							p_rowFound["process"] = this.m_strUserProcesses;
							p_rowFound.EndEdit();
							p_rowFound = null;
						}

						

					}


					if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
					if (this.btnDelete.Enabled==false) this.btnDelete.Enabled=true;
					if (this.btnEdit.Enabled==false) this.btnEdit.Enabled=true;
				}
				frmTemp.Close();
				frmTemp.Dispose();
				frmTemp=null;
			}
			catch (Exception e)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_contact_list:EditForm() \n" + 
					"Err Msg - " + e.Message,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);

				this.m_intError=-1;
			}
		}

		private void cmbFilter_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			string strFilterList="";
			string strSearchString = this.cmbFilter.Text.Trim().ToUpper();
			switch (strSearchString)
			{
				case "CORE ANALYSIS":
					strSearchString = "CORE";
					break;
				case "ALL":
					strSearchString = "";
					break;
			}
			if (strSearchString.Trim().Length > 0)
			{
				for (int x=0;x<=this.m_ado.m_DataSet.Tables["user_processes"].Rows.Count-1;x++)
				{
					string strSearch = this.m_ado.m_DataSet.Tables["user_processes"].Rows[x]["process"].ToString().Trim().ToUpper();
					if (strSearch.IndexOf(strSearchString,0) >=0)
					{
						if (strFilterList.Trim().Length == 0)
						{
							strFilterList = "'" + this.m_ado.m_DataSet.Tables["user_processes"].Rows[x]["name"].ToString().Trim().ToUpper() + "'";
						}
						else
						{
							strFilterList += ",'" + this.m_ado.m_DataSet.Tables["user_processes"].Rows[x]["name"].ToString().Trim().ToUpper() + "'";
						}


					}

				}
				if (strFilterList.Trim().Length > 0)
				{
					strSearchString = "name IN (" + strFilterList + ")";
					this.m_dv.RowFilter = strSearchString;
				}
				else
				{
					this.m_dv.RowFilter = "LEN(TRIM(name)) = 0";
				}

			}
			else
			{
				this.m_dv.RowFilter = "";

			}
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			this.m_ado.m_DataSet.Tables["contacts"].RejectChanges();
			this.btnSave.Enabled=false;
			this.ParentForm.Close();
			
		}





	}
	public class Contacts_DataGridColoredTextBoxColumn : DataGridTextBoxColumn
	{
		bool m_bEdit=false;
		FIA_Biosum_Manager.uc_contact_list uc_contact_list1;
		string m_strLastKey="";
		bool m_bNumericOnly=false;
		

		public Contacts_DataGridColoredTextBoxColumn(bool bEdit,bool bNumericOnly,FIA_Biosum_Manager.uc_contact_list p_uc)
		{
			this.m_bEdit = bEdit;
			this.m_bNumericOnly = bNumericOnly;
			this.uc_contact_list1 = p_uc;
			this.TextBox.KeyDown += new KeyEventHandler(TextBox_KeyDown);
			this.TextBox.Leave += new EventHandler(TextBox_Leave);
			this.TextBox.Enter += new EventHandler(TextBox_Enter);
		}
		

		protected override void Paint(System.Drawing.Graphics g, System.Drawing.Rectangle bounds, System.Windows.Forms.CurrencyManager source, int rowNum, System.Drawing.Brush backBrush, System.Drawing.Brush foreBrush, bool alignToRight)
		{
		  	
			// color only the columns that can be edited by the user
			try
			{
				if (this.m_bEdit == true)
				{
					backBrush = new System.Drawing.Drawing2D.LinearGradientBrush(bounds,
						Color.FromArgb(255, 200, 200), 
						Color.FromArgb(128, 20, 20),
						System.Drawing.Drawing2D.LinearGradientMode.BackwardDiagonal);
					foreBrush = new SolidBrush(Color.White);
				}
			}
			catch { /* empty catch */ }
			finally
			{
				// make sure the base class gets called to do the drawing with
				// the possibly changed brushes
				base.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight);
			}
		}
		private void TextBox_TextChanged(object sender, EventArgs e)
		{
			//MessageBox.Show("textchange");
		}
		private void TextBox_KeyDown(object sender, KeyEventArgs e)
		{
			if (this.m_bEdit == true)
			{
				
				if (this.m_bNumericOnly==true)
				{
					if (Char.IsDigit((char)e.KeyValue))
					{
						this.m_strLastKey = Convert.ToString(e.KeyValue);
						if (this.uc_contact_list1.btnSave.Enabled==false) this.uc_contact_list1.btnSave.Enabled=true;
					}
					else
					{
						if (e.KeyCode == Keys.Back)
						{
							this.m_strLastKey = Convert.ToString(e.KeyValue);
							if (this.uc_contact_list1.btnSave.Enabled==false) this.uc_contact_list1.btnSave.Enabled=true;
						}
						else
						{
							e.Handled=true;	
							SendKeys.Send("{BACKSPACE}");
						}
					}
						
				}
				else
				{
					this.m_strLastKey = Convert.ToString(e.KeyValue);
					if (this.uc_contact_list1.btnSave.Enabled==false) this.uc_contact_list1.btnSave.Enabled=true;
				}
					
				
			}
			
		}
		private void TextBox_Enter(object sender, EventArgs e)
		{
			this.m_strLastKey="";
		}
		private void TextBox_Leave(object sender, EventArgs e)
		{
			if (this.m_bEdit == true)
			{
				if (this.m_strLastKey.Trim().Length > 0)
				{
					if (this.uc_contact_list1.btnSave.Enabled==false) this.uc_contact_list1.btnSave.Enabled=true;
				}
			}
		}
		     
	}
}
