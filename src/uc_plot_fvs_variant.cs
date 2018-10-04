using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_tree_spc_conversion_and_groupings.
	/// </summary>
	public class uc_plot_fvs_variant : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.DataGrid m_dg;
		private System.Windows.Forms.GroupBox grpboxAudit;
		public System.Windows.Forms.ListView lstAudit;
		private System.Windows.Forms.Button btnEdit;
		public System.Windows.Forms.Button btnSave;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnHelp;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnAuditCheckAll;
		private System.Windows.Forms.Button btnAuditClearAll;
		private string m_strProjDir;
		private FIA_Biosum_Manager.Datasource m_DataSource;
		private string m_strPlotTable;
		private string m_strVariantTable;
		private string m_strTempMDBFile;
		private FIA_Biosum_Manager.ado_data_access m_ado;
		private string m_strConn;
		private System.Data.DataView m_dv;
		public int m_intIndex=0;
		private int m_intCurrRow=0;
		public int m_intError=0;


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
		//private string m_strDeletedList="";
		//private int m_intDeletedCount=0;
		private string m_strColumnFilterList="";
		private string m_strColumnSortList="";

        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultFvsXPSFile;

		private System.Data.DataTable m_dtTableSchema;
		private System.Windows.Forms.GroupBox grpBoxPlot;
		private System.Windows.Forms.Button btnAuditUpdate;
		private System.Windows.Forms.Button btnAuditPlotFVSVariants;



		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_plot_fvs_variant(string p_strProjDir)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

            this.m_oEnv = new env();

			// TODO: Add any initialization after the InitializeComponent call

			this.m_strProjDir = p_strProjDir;
			this.m_DataSource = new Datasource();
			m_DataSource.LoadTableColumnNamesAndDataTypes=false;
			m_DataSource.LoadTableRecordCount=false;
			m_DataSource.m_strDataSourceMDBFile = p_strProjDir.Trim() + "\\db\\project.mdb";
			m_DataSource.m_strDataSourceTableName = "datasource";
			m_DataSource.m_strScenarioId="";
			m_DataSource.populate_datasource_array();

			this.m_strPlotTable = this.m_DataSource.getValidDataSourceTableName("PLOT");
			this.m_strVariantTable = this.m_DataSource.getValidDataSourceTableName("FIADB FVS VARIANT");
			this.m_strTempMDBFile = this.m_DataSource.CreateMDBAndTableDataSourceLinks();
			if (this.m_strPlotTable.Trim().Length == 0) 
			{
				MessageBox.Show("!!Could Not Locate Plot Table!!","FVS Variant",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}
			if (this.m_strVariantTable.Trim().Length == 0)
			{
				MessageBox.Show("!!Could Not Locate The FIADB FVS Variant Table!!","FVS Variant",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}

			this.m_ado = new ado_data_access();
			this.m_strConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"","");
			this.m_ado.OpenConnection(this.m_strConn);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_ado = null;
				return ;

			}
			
			
            				 
          


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

		public void loadvalues()
		{
			string strColumnName="";
			this.InitializePopup();
				                                         
			this.m_ado.m_DataSet = new DataSet("plot");
			this.m_ado.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
			
			this.InitializeOleDbTransactionCommands();
            this.m_ado.m_strSQL = "select biosum_plot_id,statecd,countycd,plot,fvs_variant,fvsloccode,half_state from " + this.m_strPlotTable + ";";
			this.m_dtTableSchema = this.m_ado.getTableSchema(this.m_ado.m_OleDbConnection,
				                                             this.m_ado.m_OleDbTransaction,
				                                             this.m_ado.m_strSQL);
			if (this.m_ado.m_intError == 0)
			{
				this.m_ado.m_OleDbCommand = this.m_ado.m_OleDbConnection.CreateCommand();
				this.m_ado.m_OleDbCommand.CommandText = this.m_ado.m_strSQL;
				this.m_ado.m_OleDbDataAdapter.SelectCommand = this.m_ado.m_OleDbCommand;
				this.m_ado.m_OleDbDataAdapter.SelectCommand.Transaction = this.m_ado.m_OleDbTransaction;
				try 
				{

					this.m_ado.m_OleDbDataAdapter.Fill(this.m_ado.m_DataSet,"plot");

					//define the primary key
					DataColumn[] myColArray = new DataColumn[1];
					myColArray[0] = this.m_ado.m_DataSet.Tables["plot"].Columns["biosum_plot_id"];
					this.m_ado.m_DataSet.Tables["plot"].PrimaryKey = myColArray;


					this.m_dv = new DataView(this.m_ado.m_DataSet.Tables["plot"]);
				
					this.m_dv.AllowNew = false;       //cannot append new records
					this.m_dv.AllowDelete = false;    //cannot delete records
					this.m_dv.AllowEdit = true;
					this.m_dg.CaptionText = "plot";
					m_dg.BackgroundColor=frmMain.g_oGridViewBackgroundColor;
					/***********************************************************************************
					 **assign the aColumnTextColumn as type DataGridColoredTextBoxColumn object class
					 ***********************************************************************************/
					FVSVariant_DataGridColoredTextBoxColumn aColumnTextColumn ;


					/***************************************************************
					 **custom define the grid style
					 ***************************************************************/
					DataGridTableStyle tableStyle = new DataGridTableStyle();

					/***********************************************************************
					 **map the data grid table style to the scenario rx intensity dataset
					 ***********************************************************************/
					tableStyle.MappingName = "plot";
					tableStyle.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
					tableStyle.BackColor = frmMain.g_oGridViewRowBackgroundColor;
					tableStyle.ForeColor = frmMain.g_oGridViewRowForegroundColor;
					tableStyle.SelectionBackColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
   
					/******************************************************************************
					 **since the dataset has things like field name and number of columns,
					 **we will use those to create new columnstyles for the columns in our grid
					 ******************************************************************************/
					//get the number of columns from the scenario_rx_intensity data set
					int numCols = this.m_ado.m_DataSet.Tables["plot"].Columns.Count;
                
                    
					/************************************************
					 **loop through all the columns in the dataset	
					 ************************************************/
					for(int i = 0; i < numCols; ++i)
					{
						strColumnName = this.m_ado.m_DataSet.Tables["plot"].Columns[i].ColumnName;
					
						if (strColumnName.Trim().ToUpper() == "FVS_VARIANT" || strColumnName.Trim().ToUpper() == "FVSLOCCODE")
						{
							/******************************************************************
							**create a new instance of the DataGridColoredTextBoxColumn class
							******************************************************************/
							aColumnTextColumn = new FVSVariant_DataGridColoredTextBoxColumn(true,false,this);
							aColumnTextColumn.TextBox.MaxLength=2;
							aColumnTextColumn.ReadOnly=false;
						}
						else
						{
							/******************************************************************
							 **create a new instance of the DataGridColoredTextBoxColumn class
							 ******************************************************************/
							aColumnTextColumn = new FVSVariant_DataGridColoredTextBoxColumn(false,false,this);
							/***********************************
							 **all columns are read-only except
							 **the edit columns
							 ***********************************/
							aColumnTextColumn.ReadOnly=true;

						}

						

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
					if (frmMain.g_oGridViewFont != null) this.m_dg.Font = frmMain.g_oGridViewFont;

					this.m_dg.TableStyles.Clear();
					this.m_dg.TableStyles.Add(tableStyle);

					this.m_dg.DataSource = this.m_dv;  

					if (this.m_ado.m_DataSet.Tables["plot"].Rows.Count > 0)
					{
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

			
			

				if (this.m_ado.m_DataSet.Tables["plot"].Rows.Count == 0)
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
			int intLeft=0;
			int intTop=0;
			int x=0;
			int y=0;
			System.Windows.Forms.DialogResult dlgResult;
			System.Windows.Forms.CurrencyManager p_cm;
			System.Data.DataView p_dv;
			int intCurrRow=0;

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
					//---------------form----------------//
					//declare form and instatiate
					FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();

					
					strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
					if (strCol.ToUpper() == "FVS_VARIANT")
					{
						//------------text box------------//
						//instatiate numeric text class
						System.Windows.Forms.TextBox p_textbox = new TextBox();

						
						//define form properties
						frmTemp.Width = 200;
						frmTemp.Height = 200;
						frmTemp.WindowState = System.Windows.Forms.FormWindowState.Normal;
						frmTemp.MaximizeBox = false;
						frmTemp.MinimizeBox = false;
						frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
						frmTemp.Text = "Modify (" + strCol.Trim() + ")";
		                    
						//define numeric text class properties
						p_textbox.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
						p_textbox.Name = "txtModify";
						p_textbox.TabIndex = 0;
						p_textbox.Tag = "";
						p_textbox.Visible = true;
						p_textbox.Enabled = true;
						frmTemp.Controls.Add(p_textbox);  //add the text box to the form
						p_textbox.Height = 100;

						p_textbox.Top = (int)(frmTemp.ClientSize.Height * .50) - (int)(p_textbox.Height * .50);
						p_textbox.Width = frmTemp.ClientSize.Width - 20;
						p_textbox.Left = 10;
						
						p_textbox.MaxLength = 2;
						p_textbox.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
						
						p_textbox.Text = "";
						intLeft = p_textbox.Left;
						intTop = p_textbox.Top;
						frmTemp.txtBox = p_textbox;
						frmTemp.strCallingFormType="TS";
					}
					

						//-----------label---------------//
						System.Windows.Forms.Label lblTemp = new Label();
						lblTemp.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
						lblTemp.Location = new System.Drawing.Point(168, 53);
						lblTemp.Name = "lblTemp";
						lblTemp.Size = new System.Drawing.Size(128, 15);
						lblTemp.TabIndex = 1;
						lblTemp.Text = "New Value";
						frmTemp.Controls.Add(lblTemp);
						lblTemp.Top = intTop - lblTemp.Height - 10;
						lblTemp.Left = intLeft;
						lblTemp.Visible=true;


						//----------------OK button-----------------//
						System.Windows.Forms.Button btnTempOK = new Button();
						btnTempOK.Location = new System.Drawing.Point(392, 328);
						btnTempOK.Name = "btnOK";
						btnTempOK.Size = new System.Drawing.Size(80, 32);
						btnTempOK.TabIndex = 2;
						btnTempOK.Text = "OK";
						btnTempOK.Click += new System.EventHandler(frmTemp.btnOK_Click);
						frmTemp.Controls.Add(btnTempOK);
						btnTempOK.Top = intTop + btnTempOK.Height + 10;
						btnTempOK.Left = (int)(frmTemp.Width * .50)  - btnTempOK.Width;
						btnTempOK.Visible=true;

						//----------------Cancel button-----------------//
						System.Windows.Forms.Button btnTempCancel = new Button();
						btnTempCancel.Location = new System.Drawing.Point(392, 328);
						btnTempCancel.Name = "btnCancel";
						btnTempCancel.Size = new System.Drawing.Size(80, 32);
						btnTempCancel.TabIndex = 2;
						btnTempCancel.Text = "CANCEL";
						btnTempCancel.Click += new System.EventHandler(frmTemp.btnCancel_Click);
						frmTemp.Controls.Add(btnTempCancel);
						btnTempCancel.Top = intTop + btnTempCancel.Height + 10;
						btnTempCancel.Left = (int)(frmTemp.Width * .50) ;
						btnTempCancel.Visible=true;
						frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
						dlgResult = frmTemp.ShowDialog();
						if (dlgResult.ToString().Trim().ToUpper() == "OK" )
						{
							try
							{
								p_cm = (CurrencyManager)this.BindingContext[this.m_dg.DataSource,this.m_dg.DataMember];
								p_dv = (DataView)p_cm.List;
								intCurrRow= this.m_intCurrRow-1;
								y = p_dv.Count;

								/***************************************************
								 **if the column being modified has a filter or sort
								 **applied to it then the rows will change when
								 **the edit value is applied to the row.
								 **Therefore, check if the edited column is also
								 **filtered/sorted to set the variables used to handle
								 **the row changes that will occur when the edit
								 **is applied to the column.
								 *************************************************/
								int intModifyCount=0;
								bool bApplyModifyCount=false;
								string strSearch=strCol.Trim() + ",";
								if (this.m_dv.RowFilter.ToString().Trim().Length > 0)
								{
									if (this.m_strColumnFilterList.IndexOf(strSearch,0) >= 0)
										bApplyModifyCount = true;
								}
								if (bApplyModifyCount==false)
								{
									if (this.m_dv.Sort.ToString().Trim().Length > 0)
									{
										if (this.m_strColumnSortList.IndexOf(strSearch,0) >= 0)
											bApplyModifyCount = true;
									}
								}
								for (x=0; x <= y-1;x++)
								{

									if (m_dg.IsSelected(x)|| x == intCurrRow)
									{
										if (bApplyModifyCount==true)
										{
											m_dg.CurrentRowIndex = x-intModifyCount;
										}
										else
										{
											m_dg.CurrentRowIndex=x;
										}
								
										if (frmTemp.strCallingFormType=="TS")
										{
											m_dg[m_dg.CurrentRowIndex,this.m_intPopupColumn] = frmTemp.txtBox.Text ;

										}
										else if (frmTemp.strCallingFormType=="TN")
										{
											if (frmTemp.txtNumeric.Text.Trim().Length > 0)
											{
												if (frmTemp.txtNumeric.Text.Trim().Length == 1 &&
													frmTemp.txtNumeric.Text.Trim() == ".")
												{
													this.m_dg[x,this.m_intPopupColumn] = System.DBNull.Value;
												}
												else this.m_dg[x,this.m_intPopupColumn] = m_dg[x,this.m_intPopupColumn] = frmTemp.txtNumeric.Text;
											}
											else
											{
												this.m_dg[x,this.m_intPopupColumn] = System.DBNull.Value;
											}
										}
										else if (frmTemp.strCallingFormType=="TD")
										{
											if (frmTemp.txtNumeric.Text.Trim().Length > 0)
											{
												if (frmTemp.txtNumeric.Text.Trim().Length == 1 &&
													frmTemp.txtNumeric.Text.Trim() == ".")
												{
													this.m_dg[x,this.m_intPopupColumn] = System.DBNull.Value;
												}
												else this.m_dg[x,this.m_intPopupColumn] = m_dg[x,this.m_intPopupColumn] = frmTemp.txtNumeric.Text;
											}
											else
											{
												this.m_dg[x,this.m_intPopupColumn] = System.DBNull.Value;
											}
										}
										else if (frmTemp.strCallingFormType=="TM")
										{
											m_dg[x,this.m_intPopupColumn] = frmTemp.txtMoney.Text.Substring(1,frmTemp.txtMoney.Text.Trim().Length - 1);
										}
										if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
										intModifyCount++;

									}
								}
								this.m_dg.SetDataBinding(this.m_dv,"");
								this.m_dg.Update();
							
							}
							catch (Exception err)
							{
								MessageBox.Show(err.Message);
							}

						
						}
						frmTemp = null;
					break;
				case "REMOVE FILTER":
					this.m_dv.RowFilter="";
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
			//string strCount1="";
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


		private void m_dg_CurrentCellChanged(object sender, 
			System.EventArgs e)
		{
			if (this.m_intCurrRow > 0)
			{
				if (this.m_dg.CurrentRowIndex != this.m_intCurrRow - 1)
				{
					this.m_intCurrRow = this.m_dg.CurrentRowIndex + 1;
				}
			}
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

		private void m_dg_TextBox_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			switch (e.Button)
			{
				case System.Windows.Forms.MouseButtons.Right :
					this.m_mnuDataGridPopup.Show(this,new Point(e.X,e.Y));
					MessageBox.Show("here i am");
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


		public string strProjectDirectory
		{
			set
			{
				this.m_strProjDir = value;
			}
			get
			{
				return this.m_strProjDir;
			}
		}

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnHelp = new System.Windows.Forms.Button();
            this.grpBoxPlot = new System.Windows.Forms.GroupBox();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.m_dg = new System.Windows.Forms.DataGrid();
            this.grpboxAudit = new System.Windows.Forms.GroupBox();
            this.btnAuditClearAll = new System.Windows.Forms.Button();
            this.btnAuditCheckAll = new System.Windows.Forms.Button();
            this.btnAuditUpdate = new System.Windows.Forms.Button();
            this.btnAuditPlotFVSVariants = new System.Windows.Forms.Button();
            this.lstAudit = new System.Windows.Forms.ListView();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.grpBoxPlot.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.m_dg)).BeginInit();
            this.grpboxAudit.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Controls.Add(this.grpBoxPlot);
            this.groupBox1.Controls.Add(this.grpboxAudit);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(704, 616);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(592, 576);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(96, 32);
            this.btnClose.TabIndex = 47;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(8, 576);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(96, 32);
            this.btnHelp.TabIndex = 46;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // grpBoxPlot
            // 
            this.grpBoxPlot.Controls.Add(this.btnEdit);
            this.grpBoxPlot.Controls.Add(this.btnSave);
            this.grpBoxPlot.Controls.Add(this.btnCancel);
            this.grpBoxPlot.Controls.Add(this.m_dg);
            this.grpBoxPlot.Location = new System.Drawing.Point(24, 272);
            this.grpBoxPlot.Name = "grpBoxPlot";
            this.grpBoxPlot.Size = new System.Drawing.Size(664, 288);
            this.grpBoxPlot.TabIndex = 29;
            this.grpBoxPlot.TabStop = false;
            this.grpBoxPlot.Text = "Plot Table";
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(245, 240);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(64, 32);
            this.btnEdit.TabIndex = 47;
            this.btnEdit.Text = "Edit";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnSave
            // 
            this.btnSave.Enabled = false;
            this.btnSave.Location = new System.Drawing.Point(309, 240);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(64, 32);
            this.btnSave.TabIndex = 48;
            this.btnSave.Text = "Save";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(373, 240);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(64, 32);
            this.btnCancel.TabIndex = 49;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // m_dg
            // 
            this.m_dg.DataMember = "";
            this.m_dg.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.m_dg.Location = new System.Drawing.Point(8, 24);
            this.m_dg.Name = "m_dg";
            this.m_dg.Size = new System.Drawing.Size(648, 208);
            this.m_dg.TabIndex = 2;
            this.m_dg.MouseDown += new System.Windows.Forms.MouseEventHandler(this.m_dg_MouseDown);
            this.m_dg.MouseUp += new System.Windows.Forms.MouseEventHandler(this.m_dg_MouseUp);
            // 
            // grpboxAudit
            // 
            this.grpboxAudit.Controls.Add(this.btnAuditClearAll);
            this.grpboxAudit.Controls.Add(this.btnAuditCheckAll);
            this.grpboxAudit.Controls.Add(this.btnAuditUpdate);
            this.grpboxAudit.Controls.Add(this.btnAuditPlotFVSVariants);
            this.grpboxAudit.Controls.Add(this.lstAudit);
            this.grpboxAudit.Location = new System.Drawing.Point(24, 48);
            this.grpboxAudit.Name = "grpboxAudit";
            this.grpboxAudit.Size = new System.Drawing.Size(664, 216);
            this.grpboxAudit.TabIndex = 28;
            this.grpboxAudit.TabStop = false;
            this.grpboxAudit.Text = "Audit Results";
            // 
            // btnAuditClearAll
            // 
            this.btnAuditClearAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAuditClearAll.Location = new System.Drawing.Point(88, 136);
            this.btnAuditClearAll.Name = "btnAuditClearAll";
            this.btnAuditClearAll.Size = new System.Drawing.Size(72, 32);
            this.btnAuditClearAll.TabIndex = 32;
            this.btnAuditClearAll.Text = "Clear All";
            this.btnAuditClearAll.Click += new System.EventHandler(this.btnAuditClearAll_Click);
            // 
            // btnAuditCheckAll
            // 
            this.btnAuditCheckAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAuditCheckAll.Location = new System.Drawing.Point(16, 136);
            this.btnAuditCheckAll.Name = "btnAuditCheckAll";
            this.btnAuditCheckAll.Size = new System.Drawing.Size(72, 32);
            this.btnAuditCheckAll.TabIndex = 31;
            this.btnAuditCheckAll.Text = "Check All";
            this.btnAuditCheckAll.Click += new System.EventHandler(this.btnAuditCheckAll_Click);
            // 
            // btnAuditUpdate
            // 
            this.btnAuditUpdate.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAuditUpdate.Location = new System.Drawing.Point(160, 136);
            this.btnAuditUpdate.Name = "btnAuditUpdate";
            this.btnAuditUpdate.Size = new System.Drawing.Size(312, 32);
            this.btnAuditUpdate.TabIndex = 30;
            this.btnAuditUpdate.Text = "Update Plot Records With FIADB FVS Variant Table";
            this.btnAuditUpdate.Click += new System.EventHandler(this.btnAuditAdd_Click);
            // 
            // btnAuditPlotFVSVariants
            // 
            this.btnAuditPlotFVSVariants.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAuditPlotFVSVariants.Location = new System.Drawing.Point(144, 184);
            this.btnAuditPlotFVSVariants.Name = "btnAuditPlotFVSVariants";
            this.btnAuditPlotFVSVariants.Size = new System.Drawing.Size(400, 24);
            this.btnAuditPlotFVSVariants.TabIndex = 29;
            this.btnAuditPlotFVSVariants.Text = "Check For Plots Without Variant Or Location Codes";
            this.btnAuditPlotFVSVariants.Click += new System.EventHandler(this.btnAuditPlotFVSVariants_Click);
            // 
            // lstAudit
            // 
            this.lstAudit.CheckBoxes = true;
            this.lstAudit.FullRowSelect = true;
            this.lstAudit.GridLines = true;
            this.lstAudit.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstAudit.HideSelection = false;
            this.lstAudit.Location = new System.Drawing.Point(16, 24);
            this.lstAudit.MultiSelect = false;
            this.lstAudit.Name = "lstAudit";
            this.lstAudit.Size = new System.Drawing.Size(632, 104);
            this.lstAudit.TabIndex = 27;
            this.lstAudit.UseCompatibleStateImageBehavior = false;
            this.lstAudit.View = System.Windows.Forms.View.Details;
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(698, 24);
            this.lblTitle.TabIndex = 1;
            this.lblTitle.Text = "FVS Variant";
            // 
            // uc_plot_fvs_variant
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_plot_fvs_variant";
            this.Size = new System.Drawing.Size(704, 616);
            this.Resize += new System.EventHandler(this.uc_tree_spc_conversion_Resize);
            this.groupBox1.ResumeLayout(false);
            this.grpBoxPlot.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.m_dg)).EndInit();
            this.grpboxAudit.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.CleanUp();
            ((frmDialog)ParentForm).ParentControl.Enabled = true;
			this.ParentForm.Close();
		}

		private void uc_tree_spc_conversion_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
				this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
				this.btnHelp.Top = this.btnClose.Top;
				this.grpboxAudit.Width = this.groupBox1.Width  - (this.grpboxAudit.Left * 2);
				this.grpBoxPlot.Width = this.grpboxAudit.Width ;
				this.lstAudit.Width = this.grpboxAudit.Width - (this.lstAudit.Left * 2);
				this.m_dg.Width = this.grpBoxPlot.Width - (this.m_dg.Left * 2);
				this.grpBoxPlot.Height = this.btnClose.Top - this.grpBoxPlot.Top - 5;
				this.btnEdit.Top =this.grpBoxPlot.Height - this.btnEdit.Height - 2;
				this.btnSave.Top = this.btnEdit.Top;
				this.btnCancel.Top = this.btnEdit.Top;
				this.m_dg.Height = this.btnEdit.Top - this.m_dg.Top - 2;
				this.btnSave.Left = (int)(this.grpBoxPlot.Width * .50) - (int)(this.btnSave.Width * .50);
				this.btnEdit.Left = this.btnSave.Left - this.btnEdit.Width;
				this.btnCancel.Left = this.btnSave.Left + this.btnCancel.Width;
				this.btnAuditPlotFVSVariants.Left = (int)(this.grpBoxPlot.Width * .50) - (int)(this.btnAuditPlotFVSVariants.Width * .50);

			}
			catch
			{
			}
		}

		private void btnAuditPlotTreeVariantCombo_Click(object sender, System.EventArgs e)
		{
		}

		private void btnAuditCheckAll_Click(object sender, System.EventArgs e)
		{
			if (this.lstAudit.Items.Count==0) return;
			for (int x=0;x<=this.lstAudit.Items.Count-1;x++)
			{
				this.lstAudit.Items[x].Checked=true;
			}
		}

		private void btnAuditClearAll_Click(object sender, System.EventArgs e)
		{
			if (this.lstAudit.Items.Count==0) return;
			if (this.lstAudit.CheckedItems.Count==0) return;
			for (int x=0;x<=this.lstAudit.Items.Count-1;x++)
			{
				this.lstAudit.Items[x].Checked=false;
			}
		}

		private void btnAuditAdd_Click(object sender, System.EventArgs e)
		{
			if (this.lstAudit.CheckedItems.Count==0)
			{
				MessageBox.Show("!!No Items Selected!!","FVS Variant",
					System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				return;
			}
			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			if (this.btnEdit.Enabled==false) this.btnEdit.Enabled=true;
			int x;
			//string strKey;
			System.Data.DataRow  p_rowFound;
						
			string str;

			for (x=0;x<=this.lstAudit.Items.Count-1;x++)
			{
				if (this.lstAudit.Items[x].Checked==true)
				{
					str = "'" + this.lstAudit.Items[x].Text.Trim() + "'";
					//find the biosum_plot_id in the plot data set
					System.Object[] p_search = new Object[1];
					p_search[0] = this.lstAudit.Items[x].Text.Trim();
					p_rowFound = this.m_ado.m_DataSet.Tables["plot"].Rows.Find(p_search);
					if (p_rowFound != null)
					{
						p_rowFound["fvs_variant"] = this.lstAudit.Items[x].SubItems[4].Text.Trim();
						p_rowFound["fvsloccode"] = this.lstAudit.Items[x].SubItems[5].Text.Trim();
					}
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
		
		public void savevalues()
		{
			int intCurrRow;
		
            this.val_data();
			if (this.m_intError !=0) return;

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

			try
			{
				
				p_dtChanges = this.m_ado.m_DataSet.Tables["plot"].GetChanges();
								
				//check if any inserted rows
				//p_Rows = p_dtChanges.Select(null,null, DataViewRowState.Added);
				if (p_dtChanges.HasErrors)
				{
					this.m_ado.m_DataSet.Tables["plot"].RejectChanges();
					this.m_intError=-1;
				}
				else
				{
					this.m_ado.m_OleDbDataAdapter.Update(this.m_ado.m_DataSet.Tables["plot"]);
					this.m_ado.m_OleDbTransaction.Commit();
					this.m_ado.m_DataSet.Tables["plot"].AcceptChanges();
					this.InitializeOleDbTransactionCommands();
				}
			}
			catch (Exception caught)
			{
				this.m_intError=-1;
				MessageBox.Show(caught.Message);
				this.m_ado.m_DataSet.Tables["plot"].RejectChanges();
				//rollback the transaction to the original records 
				this.m_ado.m_OleDbTransaction.Rollback();
				
			}
			

			p_dtChanges=null;
			


			this.m_dg.CurrentRowIndex = intCurrRow;		
			this.btnSave.Enabled=false;	
            
			return;
		}

		private void btnSave_Click(object sender, System.EventArgs e)
		{
			this.savevalues();
		}
		private void InitializeOleDbTransactionCommands()
		{

			//initialize the transaction object with the connection
			this.m_ado.m_OleDbTransaction = this.m_ado.m_OleDbConnection.BeginTransaction();

			//declare in an sql select command the column(s) that can be updated
            this.m_ado.m_strSQL = "select fvs_variant, fvsloccode from " + this.m_strPlotTable + ";";
			this.m_ado.ConfigureDataAdapterUpdateCommand(this.m_ado.m_OleDbConnection,
				this.m_ado.m_OleDbDataAdapter,
				this.m_ado.m_OleDbTransaction,
				this.m_ado.m_strSQL,"select biosum_plot_id from " + this.m_strPlotTable,
				this.m_strPlotTable);

		}

		

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			this.CleanUp();
		     this.ParentForm.Close();		
		}
		private void CleanUp()
		{
			try
			{

				if (this.m_ado.m_OleDbDataAdapter != null)
					this.m_ado.m_OleDbDataAdapter.Dispose();

				if (this.m_dv != null)
					this.m_dv.Dispose();

				if (this.m_dg != null)
					this.m_dg.Dispose();

				this.m_ado.m_DataSet.Clear();
				this.m_ado.m_DataSet.Dispose();

				if (this.m_ado.m_OleDbConnection != null)
				{
					this.m_ado.m_OleDbConnection.Close();
					this.m_ado.m_OleDbConnection.Dispose();
					this.m_ado.m_OleDbConnection = null;
				}
				this.m_ado = null;
				this.m_DataSource = null;

			}
			catch (Exception e)
			{
				MessageBox.Show(e.Message);
			}

		}

		private void btnNew_Click(object sender, System.EventArgs e)
		{
			EditForm("New");
		}
		private void EditForm(string p_strAction)
		{
			
			FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
			frmTemp.MaximizeBox = false;
			frmTemp.BackColor = System.Drawing.SystemColors.Control;
			frmTemp.Text = "FVS: Plot FVS Variant (" + p_strAction + ")";

			FIA_Biosum_Manager.uc_plot_fvs_variant_edit  p_uc = new uc_plot_fvs_variant_edit();
			
			frmTemp.Controls.Add(p_uc);
			frmTemp.PlotFvsVariantEditUserControl=p_uc ;

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
			
            p_uc.strBiosumPlotId = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("biosum_plot_id")].ToString().Trim();
			p_uc.strStateCd = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("statecd")].ToString().Trim();
			p_uc.strVariant = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("fvs_variant")].ToString().Trim();
			p_uc.strFvsLocation = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("fvsloccode")].ToString().Trim();
			p_uc.strCountyCd = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("countycd")].ToString().Trim();
			p_uc.strHalfState = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("half_state")].ToString().Trim();
			p_uc.strPlot = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("plot")].ToString().Trim();


			
			System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
			if (result==System.Windows.Forms.DialogResult.OK)
			{
				this.m_dg[this.m_intCurrRow-1,this.getGridColumn("fvs_variant")] = p_uc.strVariant;
			    if (String.IsNullOrEmpty(p_uc.strFvsLocation))
			    {
			        this.m_dg[this.m_intCurrRow - 1, this.getGridColumn("fvsloccode")] = DBNull.Value;
			    }
			    else
			    {
			        this.m_dg[this.m_intCurrRow - 1, this.getGridColumn("fvsloccode")] = p_uc.strFvsLocation;
			    }

			    this.m_dg.SetDataBinding(this.m_dv,"");
				this.m_dg.Update();
				if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
				if (this.btnEdit.Enabled==false) this.btnEdit.Enabled=true;
			}
			frmTemp.Close();
			frmTemp.Dispose();
			frmTemp=null;
			

		}
		
		private void btnEdit_Click(object sender, System.EventArgs e)
		{
			this.EditForm("Edit");
		}
		private void val_data()
		{

		}

		private void btnAuditPlotFVSVariants_Click(object sender, System.EventArgs e)
		{
			frmMain.g_sbpInfo.Text = "Running Audit...Stand By";
			int intPlotMissingVariantFoundInMasterVariantTableCount=0;
			int intPlotCount=0;
			int intPlotCountWithVariant=0;
			int intJoinCount=0;
			this.lstAudit.Clear();
			string strMsg="";
			//first get unique tree species

			
			this.lstAudit.Columns.Add("biosum_plot_id", 125, HorizontalAlignment.Left);
			this.lstAudit.Columns.Add("statecd", 80, HorizontalAlignment.Left);
			this.lstAudit.Columns.Add("countycd", 80, HorizontalAlignment.Left);
			this.lstAudit.Columns.Add("plot", 80, HorizontalAlignment.Left);
			this.lstAudit.Columns.Add("fvs_variant", 80, HorizontalAlignment.Left);
			this.lstAudit.Columns.Add("fvsloccode", 80, HorizontalAlignment.Left);

			//get the total plot record count
			this.m_ado.m_strSQL = "select count(*) from " + this.m_strPlotTable + " p where len(trim(p.biosum_plot_id)) > 0 AND mid(p.biosum_plot_id,1,1)='1'";
			intPlotCount=Convert.ToInt32(m_ado.getRecordCount(this.m_ado.m_OleDbConnection,this.m_ado.m_OleDbTransaction,m_ado.m_strSQL,this.m_strPlotTable));

			//get the total plot record count
			this.m_ado.m_strSQL = "select count(*) from " + this.m_strPlotTable + " p where len(trim(p.biosum_plot_id)) > 0 AND mid(p.biosum_plot_id,1,1)='1' AND p.fvs_variant IS NOT NULL AND  LEN(TRIM(p.fvs_variant))>0";
			intPlotCountWithVariant=Convert.ToInt32(m_ado.getRecordCount(this.m_ado.m_OleDbConnection,this.m_ado.m_OleDbTransaction,m_ado.m_strSQL,this.m_strPlotTable));
			
            //join plot and fiadb_fvs_variant table. ensure no duplicates by using DISTINCT 
            this.m_ado.m_strSQL = "SELECT COUNT(*) AS reccount " +
                                  "FROM " + this.m_strPlotTable + " p," + 
                                  "(SELECT DISTINCT statecd,countycd,plot FROM " + m_strVariantTable + ") v " +
                                  "WHERE len(trim(p.biosum_plot_id)) > 0 AND " +
                                        "mid(p.biosum_plot_id,1,1)='1' AND " +
                                        "v.statecd=p.statecd AND " +
                                        "v.countycd=p.countycd AND " +
                                        "v.plot=p.plot";

			intJoinCount = (int)m_ado.getSingleDoubleValueFromSQLQuery(m_ado.m_OleDbConnection,m_ado.m_OleDbTransaction, m_ado.m_strSQL, "join");
            
			this.m_ado.m_strSQL = "select p.biosum_plot_id,p.statecd,p.countycd,p.plot,v.fvs_variant, v.fvsloccode " + 
				"from " + this.m_strVariantTable + " v," + 
				this.m_strPlotTable + " p " + 
				"where LEN(TRIM(p.biosum_plot_id)) > 0 AND mid(p.biosum_plot_id,1,1)='1' and " + 
				"v.statecd=p.statecd and " + 
				"v.countycd=p.countycd and " + 
				"v.plot=p.plot and " + 
				"(trim(ucase(v.fvs_variant)) <> trim(ucase(p.fvs_variant)) or " + 
				"len(trim(v.fvs_variant)) > 0 and (p.fvs_variant is null OR len(trim(p.fvs_variant))=0));";


			this.m_ado.SqlQueryReader(this.m_ado.m_OleDbConnection,this.m_ado.m_OleDbTransaction,this.m_ado.m_strSQL);
			if (this.m_ado.m_intError==0)
			{
				if (this.m_ado.m_OleDbDataReader.HasRows)
				{
					while (this.m_ado.m_OleDbDataReader.Read())
					{
					    if (this.m_ado.m_OleDbDataReader["statecd"] != System.DBNull.Value &&
					        this.m_ado.m_OleDbDataReader["countycd"] != System.DBNull.Value &&
					        this.m_ado.m_OleDbDataReader["plot"] != System.DBNull.Value &&
					        this.m_ado.m_OleDbDataReader["fvs_variant"] != System.DBNull.Value)
						{
							this.lstAudit.Items.Add(this.m_ado.m_OleDbDataReader["biosum_plot_id"].ToString().Trim());
							this.lstAudit.Items[this.lstAudit.Items.Count-1].SubItems.Add(this.m_ado.m_OleDbDataReader["statecd"].ToString().Trim());
							this.lstAudit.Items[this.lstAudit.Items.Count-1].SubItems.Add(this.m_ado.m_OleDbDataReader["countycd"].ToString().Trim());
							this.lstAudit.Items[this.lstAudit.Items.Count-1].SubItems.Add(this.m_ado.m_OleDbDataReader["plot"].ToString().Trim());              
							this.lstAudit.Items[this.lstAudit.Items.Count-1].SubItems.Add(this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim());
							this.lstAudit.Items[this.lstAudit.Items.Count-1].SubItems.Add(this.m_ado.m_OleDbDataReader["fvsloccode"].ToString().Trim());
							intPlotMissingVariantFoundInMasterVariantTableCount++;
						}
					}
				}
				else
				{
					intPlotMissingVariantFoundInMasterVariantTableCount=0;
				}
				this.m_ado.m_OleDbDataReader.Close();
				frmMain.g_sbpInfo.Text = "Ready";
				if (intPlotCount == intPlotCountWithVariant)
				{
					strMsg="Audit Passed. \r\n\r\n Every plot has an FVS variant assignment.";
				}
				else
				{
					if (intPlotCount-intPlotCountWithVariant > 1)
					   strMsg="Audit Failed.\r\n\r\n" + Convert.ToString(intPlotCount-intPlotCountWithVariant).Trim() + " plots do not have FVS variant and/or location code assignments.";
					else
					   strMsg="Audit Failed.\r\n\r\n" + Convert.ToString(intPlotCount-intPlotCountWithVariant).Trim() + " plot does not have an FVS variant and/or location code assignment.";
				}
				if (intPlotMissingVariantFoundInMasterVariantTableCount > 0)
				{
					strMsg+="\r\n\r\n" + Convert.ToString(intPlotMissingVariantFoundInMasterVariantTableCount).Trim() + " of the project plots were found in the master plot variant table " + this.m_strVariantTable + ". \r\n" ;
				    strMsg+="The plot/variant/location code assignments are listed at the top of the form.\r\n"; 
					strMsg+="To update the plots in the project with these plot/variant/location code combinations\r\n";
					strMsg+="select the <Update Plot Records With FIADB FVS Variant Table> button.";
				}
				if (intPlotCount != intJoinCount)
				{
					strMsg+="\r\n\r\n" + "Additional Information" + "\r\n";
					strMsg+="------------------------------\r\n";
					strMsg+=Convert.ToString(intPlotCount - intJoinCount) + " StateCd + CountyCd + Plot + Variant Combination(s) were NOT FOUND in the " + this.m_strVariantTable + " table.\r\n\r\n";
                    strMsg += "You may want to ask your Biosum administrator to update the " + this.m_strVariantTable + " table with plot/variant assignments. By using the " + this.m_strVariantTable + " table,\r\n";
					strMsg+="FIA Biosum will automatically populate your project plots with the appropriate plot/variant/location code assignments.";
				}
				MessageBox.Show(strMsg,
					"FIA Biosum", 
					System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Information);

			}

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

        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "FVS", "PLOT_FVS_VARIANTS" });
        }


	}
	public class FVSVariant_DataGridColoredTextBoxColumn : DataGridTextBoxColumn
	{
		bool m_bEdit=false;
		FIA_Biosum_Manager.uc_plot_fvs_variant uc_plot_fvs_variant1;
		string m_strLastKey="";
		bool m_bNumericOnly=false;
		

		public FVSVariant_DataGridColoredTextBoxColumn(bool bEdit,bool bNumericOnly,FIA_Biosum_Manager.uc_plot_fvs_variant p_uc)
		{
			this.m_bEdit = bEdit;
			this.m_bNumericOnly = bNumericOnly;
			this.uc_plot_fvs_variant1 = p_uc;
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
				//if (this.m_strLastKey.Trim().Length == 0) 
				//{
					if (this.m_bNumericOnly==true)
					{
						if (Char.IsDigit((char)e.KeyValue))
						{
							this.m_strLastKey = Convert.ToString(e.KeyValue);
							if (this.uc_plot_fvs_variant1.btnSave.Enabled==false) this.uc_plot_fvs_variant1.btnSave.Enabled=true;
						}
						else
						{
							if (e.KeyCode == Keys.Back)
							{
								this.m_strLastKey = Convert.ToString(e.KeyValue);
								if (this.uc_plot_fvs_variant1.btnSave.Enabled==false) this.uc_plot_fvs_variant1.btnSave.Enabled=true;
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
						if (this.uc_plot_fvs_variant1.btnSave.Enabled==false) this.uc_plot_fvs_variant1.btnSave.Enabled=true;
					}
					
				//}
			}
			//if (this.m_bEdit==true) this.uc_gridview1.btnSave.Enabled=true;
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
					if (this.uc_plot_fvs_variant1.btnSave.Enabled==false) this.uc_plot_fvs_variant1.btnSave.Enabled=true;
				}
			}
		}
		     
	}
}
