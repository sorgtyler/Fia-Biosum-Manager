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
	public class uc_fvs_tree_spc_conversion : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.DataGrid m_dg;
		private System.Windows.Forms.GroupBox grpboxAudit;
		private System.Windows.Forms.Button btnAuditPlotTreeVariantCombo;
		public System.Windows.Forms.ListView lstAudit;
		private System.Windows.Forms.GroupBox grpBoxTreeSpc;
		public System.Windows.Forms.Button btnDelete;
		private System.Windows.Forms.Button btnNew;
		private System.Windows.Forms.Button btnEdit;
		public System.Windows.Forms.Button btnSave;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnHelp;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnAuditAdd;
		private System.Windows.Forms.Button btnAuditCheckAll;
		private System.Windows.Forms.Button btnAuditClearAll;
		private string m_strProjDir;
		private FIA_Biosum_Manager.Datasource m_DataSource;
		private string m_strPlotTable;
		private string m_strTreeTable;
		private string m_strTreeSpcCvtTable;
		private string m_strTreeSpcTable;
		private string m_strCondTable;
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
		private string m_strDeletedList="";
		private int m_intDeletedCount=0;
		private string m_strColumnFilterList="";
		private string m_strColumnSortList="";

		private System.Data.DataTable m_dtTableSchema;



		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_fvs_tree_spc_conversion(string p_strProjDir)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call

			this.m_strProjDir = p_strProjDir;
			this.m_DataSource = new Datasource(p_strProjDir);
			this.m_strPlotTable = this.m_DataSource.getValidDataSourceTableName("PLOT");
			this.m_strTreeTable = this.m_DataSource.getValidDataSourceTableName("TREE");
			this.m_strTreeSpcCvtTable = this.m_DataSource.getValidDataSourceTableName("TREE SPECIES CONVERSION");
			this.m_strTreeSpcTable = this.m_DataSource.getValidDataSourceTableName("TREE SPECIES");
			this.m_strCondTable = this.m_DataSource.getValidDataSourceTableName("CONDITION");
			this.m_strTempMDBFile = this.m_DataSource.CreateMDBAndTableDataSourceLinks();
			if (this.m_strPlotTable.Trim().Length == 0) 
			{
				MessageBox.Show("!!Could Not Locate Plot Table!!","Tree Species Conversion",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}
			if (this.m_strTreeTable.Trim().Length == 0)
			{
				MessageBox.Show("!!Could Not Locate Tree Table!!","Tree Species Conversion",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}
			if (this.m_strTreeSpcCvtTable.Trim().Length == 0)
			{
				MessageBox.Show("!!Could Not Locate Tree Species Conversion Table!!","Tree Species Conversion",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}
			if (this.m_strTreeSpcTable.Trim().Length == 0)
			{
				MessageBox.Show("!!Could Not Locate Tree Species Table!!","Tree Species Conversion",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
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
				                                         
			this.m_ado.m_DataSet = new DataSet("tree_species_conversion");
			this.m_ado.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
			
			this.InitializeOleDbTransactionCommands();
            this.m_ado.m_strSQL = "select id, fvs_variant, fia_spcd, common_name, convert_to_fia_spcd,comments from " + this.m_strTreeSpcCvtTable + " order by fvs_variant, fia_spcd;";
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

					this.m_ado.m_OleDbDataAdapter.Fill(this.m_ado.m_DataSet,"tree_species_conversion");
					this.m_dv = new DataView(this.m_ado.m_DataSet.Tables["tree_species_conversion"]);
				
					this.m_dv.AllowNew = false;       //cannot append new records
					this.m_dv.AllowDelete = false;    //cannot delete records
					this.m_dv.AllowEdit = true;
					this.m_dg.CaptionText = "tree_species_conversion";
					m_dg.BackgroundColor=frmMain.g_oGridViewBackgroundColor;
					/***********************************************************************************
					 **assign the aColumnTextColumn as type DataGridColoredTextBoxColumn object class
					 ***********************************************************************************/
					TreeSpcCvt_DataGridColoredTextBoxColumn aColumnTextColumn ;


					/***************************************************************
					 **custom define the grid style
					 ***************************************************************/
					DataGridTableStyle tableStyle = new DataGridTableStyle();

					/***********************************************************************
					 **map the data grid table style to the scenario rx intensity dataset
					 ***********************************************************************/
					tableStyle.MappingName = "tree_species_conversion";
   
					/******************************************************************************
					 **since the dataset has things like field name and number of columns,
					 **we will use those to create new columnstyles for the columns in our grid
					 ******************************************************************************/
					//get the number of columns from the scenario_rx_intensity data set
					int numCols = this.m_ado.m_DataSet.Tables["tree_species_conversion"].Columns.Count;
                
                    
					/************************************************
					 **loop through all the columns in the dataset	
					 ************************************************/
					for(int i = 0; i < numCols; ++i)
					{
						strColumnName = this.m_ado.m_DataSet.Tables["tree_species_conversion"].Columns[i].ColumnName;
					
						if (strColumnName.Trim().ToUpper() == "CONVERT_TO_FIA_SPCD")
						{
							/******************************************************************
							**create a new instance of the DataGridColoredTextBoxColumn class
							******************************************************************/
							aColumnTextColumn = new TreeSpcCvt_DataGridColoredTextBoxColumn(true,true,this);
							aColumnTextColumn.TextBox.MaxLength=3;
							aColumnTextColumn.ReadOnly=false;
						}
						else if (strColumnName.Trim().ToUpper() == "COMMENTS" || strColumnName.Trim().ToUpper() == "COMMON_NAME")
						{
							aColumnTextColumn = new TreeSpcCvt_DataGridColoredTextBoxColumn(true,false,this);
							aColumnTextColumn.TextBox.MaxLength=50;
							aColumnTextColumn.ReadOnly=false;
						}
						else
						{
							/******************************************************************
							 **create a new instance of the DataGridColoredTextBoxColumn class
							 ******************************************************************/
							aColumnTextColumn = new TreeSpcCvt_DataGridColoredTextBoxColumn(false,false,this);
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
					this.m_dg.TableStyles.Clear();
					this.m_dg.TableStyles.Add(tableStyle);

					this.m_dg.DataSource = this.m_dv;  

					if (this.m_ado.m_DataSet.Tables["tree_species_conversion"].Rows.Count > 0)
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

			
			

				if (this.m_ado.m_DataSet.Tables["tree_species_conversion"].Rows.Count == 0)
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
					if (strCol.ToUpper()== "CONVERT_TO_FIA_SPCD")
					{

						//------------text box------------//
						//instatiate numeric text class
						FIA_Biosum_Manager.txtNumeric p_txtDefault = new FIA_Biosum_Manager.txtNumeric(3,0);

						
						//define form properties
						frmTemp.Width = 200;
						frmTemp.Height = 200;
						frmTemp.WindowState = System.Windows.Forms.FormWindowState.Normal;
						frmTemp.MaximizeBox = false;
						frmTemp.MinimizeBox = false;
						frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
						frmTemp.Text = "Modify";
		                    
						//define numeric text class properties
						p_txtDefault.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
						p_txtDefault.Name = "txtModify";
						p_txtDefault.TabIndex = 0;
						p_txtDefault.Tag = "";
						p_txtDefault.Visible = true;
						p_txtDefault.Enabled = true;
						frmTemp.Controls.Add(p_txtDefault);  //add the text box to the form
						p_txtDefault.Height = 100;

						p_txtDefault.Top = (int)(frmTemp.ClientSize.Height * .50) - (int)(p_txtDefault.Height * .50);
						p_txtDefault.Width = frmTemp.ClientSize.Width - 20;
						p_txtDefault.Left = 10;

						p_txtDefault.bEdit=true;
						p_txtDefault.ReadOnly=false;
						p_txtDefault.Text = "";
						intLeft = p_txtDefault.Left;
						intTop = p_txtDefault.Top;
						frmTemp.txtNumeric = p_txtDefault;
						frmTemp.strCallingFormType="TN";
					}
					else if (strCol.ToUpper() == "COMMENTS" || strCol.ToUpper() == "COMMON_NAME")
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
						frmTemp.Text = "Modify";
		                    
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
						
						p_textbox.MaxLength = 50;
						
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
							this.m_strDeletedList += "," +  p_dv[x]["id"].ToString().Trim();
						}
						else
						{
							this.m_strDeletedList = p_dv[x]["id"].ToString().Trim();
						}
						strList += "," + p_dv[x]["id"].ToString().Trim() + ",";
						this.m_intDeletedCount++;
						
					}
				}
			if (this.m_intDeletedCount > 0)
			{
				string strMsg = this.m_intDeletedCount.ToString().Trim() + " record(s) will be deleted from the table. \n Do you want to delete the record(s) from the table? (Y/N)";
				DialogResult result = MessageBox.Show(strMsg,"Delete Selected Tree Species", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
				if (result == DialogResult.Yes)
				{
					this.m_dv.AllowDelete = true;

					if (this.m_intError==0 && this.m_ado.m_intError==0)
					{
						//delete any plots that are in the delete list
						for (x=0; x < p_dv.Count;x++)
						{
							strValue = "," + p_dv[x]["id"].ToString().Trim() + ",";
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
					this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=true;
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
						if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled==true)
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=false;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=true;

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
			this.grpBoxTreeSpc = new System.Windows.Forms.GroupBox();
			this.btnDelete = new System.Windows.Forms.Button();
			this.btnNew = new System.Windows.Forms.Button();
			this.btnEdit = new System.Windows.Forms.Button();
			this.btnSave = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.m_dg = new System.Windows.Forms.DataGrid();
			this.grpboxAudit = new System.Windows.Forms.GroupBox();
			this.btnAuditClearAll = new System.Windows.Forms.Button();
			this.btnAuditCheckAll = new System.Windows.Forms.Button();
			this.btnAuditAdd = new System.Windows.Forms.Button();
			this.btnAuditPlotTreeVariantCombo = new System.Windows.Forms.Button();
			this.lstAudit = new System.Windows.Forms.ListView();
			this.lblTitle = new System.Windows.Forms.Label();
			this.groupBox1.SuspendLayout();
			this.grpBoxTreeSpc.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.m_dg)).BeginInit();
			this.grpboxAudit.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btnClose);
			this.groupBox1.Controls.Add(this.btnHelp);
			this.groupBox1.Controls.Add(this.grpBoxTreeSpc);
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
			this.btnHelp.Location = new System.Drawing.Point(8, 576);
			this.btnHelp.Name = "btnHelp";
			this.btnHelp.Size = new System.Drawing.Size(96, 32);
			this.btnHelp.TabIndex = 46;
			this.btnHelp.Text = "Help";
			// 
			// grpBoxTreeSpc
			// 
			this.grpBoxTreeSpc.Controls.Add(this.btnDelete);
			this.grpBoxTreeSpc.Controls.Add(this.btnNew);
			this.grpBoxTreeSpc.Controls.Add(this.btnEdit);
			this.grpBoxTreeSpc.Controls.Add(this.btnSave);
			this.grpBoxTreeSpc.Controls.Add(this.btnCancel);
			this.grpBoxTreeSpc.Controls.Add(this.m_dg);
			this.grpBoxTreeSpc.Location = new System.Drawing.Point(24, 272);
			this.grpBoxTreeSpc.Name = "grpBoxTreeSpc";
			this.grpBoxTreeSpc.Size = new System.Drawing.Size(664, 288);
			this.grpBoxTreeSpc.TabIndex = 29;
			this.grpBoxTreeSpc.TabStop = false;
			this.grpBoxTreeSpc.Text = "Tree Species Conversion Table";
			// 
			// btnDelete
			// 
			this.btnDelete.Enabled = false;
			this.btnDelete.Location = new System.Drawing.Point(328, 240);
			this.btnDelete.Name = "btnDelete";
			this.btnDelete.Size = new System.Drawing.Size(64, 32);
			this.btnDelete.TabIndex = 51;
			this.btnDelete.Text = "Delete";
			this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
			// 
			// btnNew
			// 
			this.btnNew.Location = new System.Drawing.Point(200, 240);
			this.btnNew.Name = "btnNew";
			this.btnNew.Size = new System.Drawing.Size(64, 32);
			this.btnNew.TabIndex = 46;
			this.btnNew.Text = "New";
			this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
			// 
			// btnEdit
			// 
			this.btnEdit.Location = new System.Drawing.Point(264, 240);
			this.btnEdit.Name = "btnEdit";
			this.btnEdit.Size = new System.Drawing.Size(64, 32);
			this.btnEdit.TabIndex = 47;
			this.btnEdit.Text = "Edit";
			this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
			// 
			// btnSave
			// 
			this.btnSave.Enabled = false;
			this.btnSave.Location = new System.Drawing.Point(392, 240);
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(64, 32);
			this.btnSave.TabIndex = 48;
			this.btnSave.Text = "Save";
			this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
			// 
			// btnCancel
			// 
			this.btnCancel.Location = new System.Drawing.Point(456, 240);
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
			this.grpboxAudit.Controls.Add(this.btnAuditAdd);
			this.grpboxAudit.Controls.Add(this.btnAuditPlotTreeVariantCombo);
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
			this.btnAuditClearAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnAuditClearAll.Location = new System.Drawing.Point(88, 136);
			this.btnAuditClearAll.Name = "btnAuditClearAll";
			this.btnAuditClearAll.Size = new System.Drawing.Size(72, 32);
			this.btnAuditClearAll.TabIndex = 32;
			this.btnAuditClearAll.Text = "Clear All";
			this.btnAuditClearAll.Click += new System.EventHandler(this.btnAuditClearAll_Click);
			// 
			// btnAuditCheckAll
			// 
			this.btnAuditCheckAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnAuditCheckAll.Location = new System.Drawing.Point(16, 136);
			this.btnAuditCheckAll.Name = "btnAuditCheckAll";
			this.btnAuditCheckAll.Size = new System.Drawing.Size(72, 32);
			this.btnAuditCheckAll.TabIndex = 31;
			this.btnAuditCheckAll.Text = "Check All";
			this.btnAuditCheckAll.Click += new System.EventHandler(this.btnAuditCheckAll_Click);
			// 
			// btnAuditAdd
			// 
			this.btnAuditAdd.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnAuditAdd.Location = new System.Drawing.Point(160, 136);
			this.btnAuditAdd.Name = "btnAuditAdd";
			this.btnAuditAdd.Size = new System.Drawing.Size(312, 32);
			this.btnAuditAdd.TabIndex = 30;
			this.btnAuditAdd.Text = "Add Checked Items To Tree Species Conversion Table";
			this.btnAuditAdd.Click += new System.EventHandler(this.btnAuditAdd_Click);
			// 
			// btnAuditPlotTreeVariantCombo
			// 
			this.btnAuditPlotTreeVariantCombo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnAuditPlotTreeVariantCombo.Location = new System.Drawing.Point(16, 184);
			this.btnAuditPlotTreeVariantCombo.Name = "btnAuditPlotTreeVariantCombo";
			this.btnAuditPlotTreeVariantCombo.Size = new System.Drawing.Size(624, 24);
			this.btnAuditPlotTreeVariantCombo.TabIndex = 29;
			this.btnAuditPlotTreeVariantCombo.Text = "FVS Input Tables: List Tree Species And Plot Variant Combinations Not Found In Tr" +
				"ee Species Conversion Table";
			this.btnAuditPlotTreeVariantCombo.Click += new System.EventHandler(this.btnAuditPlotTreeVariantCombo_Click);
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
			this.lstAudit.View = System.Windows.Forms.View.Details;
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(698, 24);
			this.lblTitle.TabIndex = 1;
			this.lblTitle.Text = "FVS Tree Species Conversion";
			// 
			// uc_fvs_tree_spc_conversion
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_fvs_tree_spc_conversion";
			this.Size = new System.Drawing.Size(704, 616);
			this.Resize += new System.EventHandler(this.uc_tree_spc_conversion_Resize);
			this.groupBox1.ResumeLayout(false);
			this.grpBoxTreeSpc.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.m_dg)).EndInit();
			this.grpboxAudit.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.CleanUp();
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
				this.grpBoxTreeSpc.Width = this.grpboxAudit.Width ;
				this.lstAudit.Width = this.grpboxAudit.Width - (this.lstAudit.Left * 2);
				this.btnAuditPlotTreeVariantCombo.Width = this.lstAudit.Width;
				this.m_dg.Width = this.grpBoxTreeSpc.Width - (this.m_dg.Left * 2);
				this.grpBoxTreeSpc.Height = this.btnClose.Top - this.grpBoxTreeSpc.Top - 5;
				this.btnNew.Top =this.grpBoxTreeSpc.Height - this.btnNew.Height - 2;
				this.btnEdit.Top = this.btnNew.Top;
				this.btnSave.Top = this.btnNew.Top;
				this.btnDelete.Top = this.btnNew.Top;
				this.btnCancel.Top = this.btnNew.Top;
				this.m_dg.Height = this.btnNew.Top - this.m_dg.Top - 2;
				this.btnDelete.Left = (int)(this.grpBoxTreeSpc.Width * .50) - (int)(this.btnDelete.Width * .50);
				this.btnEdit.Left = this.btnDelete.Left - this.btnEdit.Width;
				this.btnNew.Left = this.btnEdit.Left - this.btnNew.Width;
				this.btnSave.Left = this.btnDelete.Left + this.btnSave.Width;
				this.btnCancel.Left = this.btnSave.Left + this.btnCancel.Width;

			}
			catch
			{
			}
		}

		private void btnAuditPlotTreeVariantCombo_Click(object sender, System.EventArgs e)
		{
			this.lstAudit.Clear();
			//first get unique tree species
			string[,] strValues = new string[1000,2];

			int count=0;
			string strSpCd="";
			string strVar="";
			string strBuild="";
			string strConcat="";
			this.lstAudit.Columns.Add("fvs_variant", 80, HorizontalAlignment.Left);
			this.lstAudit.Columns.Add("spcd", 80, HorizontalAlignment.Left);

			this.m_ado.m_strSQL = "select t.spcd,p.fvs_variant " + 
				"from " + this.m_strTreeTable + " t," + 
				this.m_strPlotTable + " p " + 
				"inner join " + this.m_strCondTable + " c "  + 
				"on p.biosum_plot_id=c.biosum_plot_id " +
				"where t.biosum_cond_id=c.biosum_cond_id AND p.fvs_variant IS NOT NULL AND len(trim(p.fvs_variant)) > 0 " + 
				"order by p.fvs_variant,t.spcd;";


			this.m_ado.SqlQueryReader(this.m_ado.m_OleDbConnection,this.m_ado.m_OleDbTransaction,this.m_ado.m_strSQL);
			if (this.m_ado.m_intError==0)
			{
				if (this.m_ado.m_OleDbDataReader.HasRows)
				{
					while (this.m_ado.m_OleDbDataReader.Read())
					{
						
						if (this.m_ado.m_OleDbDataReader["spcd"] != System.DBNull.Value)
						{

							strVar="";
							strSpCd=this.m_ado.m_OleDbDataReader["spcd"].ToString().Trim();
							if (this.m_ado.m_OleDbDataReader["fvs_variant"] != System.DBNull.Value)
								strVar = this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim();

							strConcat = strSpCd + strVar;
							if (strBuild.IndexOf(strConcat,0,strBuild.Length) == -1)
							{
								strValues[count,0] = strSpCd;
								strValues[count,1] = strVar;
								strBuild = strBuild + "'" + strSpCd + strVar + "',";
								count++;
							}
						}
					}
					
					if (count > 0)
					{   
						//int y=0;
						for (int x=0; x<=count-1;x++)
						{   

							
							System.Data.DataRow[] p_rows = this.m_ado.m_DataSet.Tables["tree_species_conversion"].Select("fia_spcd = " + strValues[x,0].ToString().Trim()  +  " and trim(fvs_variant) = '" + strValues[x,1].ToString().Trim() + "'");
							if (p_rows != null)
							{
								if (p_rows.Length == 0)
								{
									this.lstAudit.Items.Add(strValues[x,1].ToString().Trim());
									this.lstAudit.Items[this.lstAudit.Items.Count-1].SubItems.Add(strValues[x,0].ToString().Trim());
								}
							}
							else
							{
								this.lstAudit.Items.Add(strValues[x,1].ToString().Trim());
								this.lstAudit.Items[this.lstAudit.Items.Count-1].SubItems.Add(strValues[x,0].ToString().Trim());
							}
						}
					}
				}
				this.m_ado.m_OleDbDataReader.Close();
			}
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
				MessageBox.Show("!!No Items Selected!!","Tree Species Conversion",
					System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				return;
			}
			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			if (this.btnDelete.Enabled==false) this.btnDelete.Enabled=true;
			if (this.btnEdit.Enabled==false) this.btnEdit.Enabled=true;
			int x;
						
			this.m_dv.AllowNew = true;
			string str;

			int intId = this.getUniqueId();
			
			for (x=0;x<=this.lstAudit.Items.Count-1;x++)
			{
				if (this.lstAudit.Items[x].Checked==true)
				{
					this.m_ado.m_strSQL = "SELECT DISTINCT common_name " + 
						"FROM " + this.m_strTreeSpcTable + " t " + 
						"WHERE SPCD = " + this.lstAudit.Items[x].SubItems[1].Text.Trim() + ";";
						                  
					str = this.m_ado.getSingleStringValueFromSQLQuery(this.m_ado.m_OleDbConnection,
						this.m_ado.m_OleDbTransaction,
						this.m_ado.m_strSQL,
						"tree_species");
					System.Data.DataRow p_row =	this.m_ado.m_DataSet.Tables["tree_species_conversion"].NewRow();
					p_row["id"] = intId;
					p_row["fvs_variant"] = this.lstAudit.Items[x].Text;
					p_row["fia_spcd"] = Convert.ToInt32(this.lstAudit.Items[x].SubItems[1].Text);
					p_row["common_name"] = str;
					p_row["convert_to_fia_spcd"] = p_row["fia_spcd"];

					this.m_ado.m_DataSet.Tables["tree_species_conversion"].Rows.Add(p_row);
					p_row=null;
					this.lstAudit.Items[x].Remove();
					intId++;
					x--;
				}
				

				this.m_dv.AllowNew = false;
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
						if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled==true)
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=false;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=true;

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
								this.m_mnuDataGridPopup.MenuItems[MENU_MODIFY].Enabled=false;
								this.m_mnuDataGridPopup.MenuItems[MENU_COUNTBYVALUE].Enabled=false;
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=false;

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
							if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled==true)
								this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=false;
							if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled==false)
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=true;

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
						this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=false;						this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled=false;
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
				
				p_dtChanges = this.m_ado.m_DataSet.Tables["tree_species_conversion"].GetChanges();
								
				//check if any inserted rows
				//p_Rows = p_dtChanges.Select(null,null, DataViewRowState.Added);
				if (p_dtChanges.HasErrors)
				{
					this.m_ado.m_DataSet.Tables["tree_species_conversion"].RejectChanges();
					this.m_intError=-1;
				}
				else
				{
					this.m_ado.m_OleDbDataAdapter.Update(this.m_ado.m_DataSet.Tables["tree_species_conversion"]);
					this.m_ado.m_OleDbTransaction.Commit();
					this.m_ado.m_DataSet.Tables["tree_species_conversion"].AcceptChanges();
					this.InitializeOleDbTransactionCommands();
				}
					
					
				

				
			}
			catch (Exception caught)
			{
				this.m_intError=-1;
				MessageBox.Show(caught.Message);
				this.m_ado.m_DataSet.Tables["tree_species_conversion"].RejectChanges();
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
			this.m_ado.m_strSQL = "select id, fvs_variant, fia_spcd, common_name, convert_to_fia_spcd,comments from " + this.m_strTreeSpcCvtTable + " order by fvs_variant, fia_spcd;";
			//initialize the transaction object with the connection
			this.m_ado.m_OleDbTransaction = this.m_ado.m_OleDbConnection.BeginTransaction();

			this.m_ado.ConfigureDataAdapterInsertCommand(this.m_ado.m_OleDbConnection,
				this.m_ado.m_OleDbDataAdapter,
				this.m_ado.m_OleDbTransaction,
				this.m_ado.m_strSQL,
				this.m_strTreeSpcCvtTable);

            this.m_ado.m_strSQL = "select fvs_variant, fia_spcd, common_name, convert_to_fia_spcd,comments from " + this.m_strTreeSpcCvtTable + " order by fvs_variant, fia_spcd;";
			this.m_ado.ConfigureDataAdapterUpdateCommand(this.m_ado.m_OleDbConnection,
				this.m_ado.m_OleDbDataAdapter,
				this.m_ado.m_OleDbTransaction,
				this.m_ado.m_strSQL,"select id from " + this.m_strTreeSpcCvtTable,
				this.m_strTreeSpcCvtTable);

			this.m_ado.m_strSQL = "select fvs_variant, fia_spcd, common_name, convert_to_fia_spcd,comments from " + this.m_strTreeSpcCvtTable + " order by fvs_variant, fia_spcd;";
			this.m_ado.ConfigureDataAdapterDeleteCommand(this.m_ado.m_OleDbConnection,
				this.m_ado.m_OleDbDataAdapter,
				this.m_ado.m_OleDbTransaction,
				"select id from " + this.m_strTreeSpcCvtTable,
				this.m_strTreeSpcCvtTable);

			


		}

		private void btnDelete_Click(object sender, System.EventArgs e)
		{
			this.DeleteRecords();
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
			frmTemp.Text = "FVS: Tree Species Conversion (" + p_strAction + ")";

			FIA_Biosum_Manager.uc_fvs_tree_spc_conversion_edit  p_uc = new uc_fvs_tree_spc_conversion_edit();
			
			frmTemp.Controls.Add(p_uc);
			frmTemp.TreeSpeciesConversionEditUserControl=p_uc ;

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
			if (p_strAction.Trim().ToUpper() == "NEW")
			{
				p_uc.strId = Convert.ToString(this.getUniqueId());
			}
			else
			{
                p_uc.strId = this.m_dg[this.m_intCurrRow-1,0].ToString().Trim();
				p_uc.strSpCd = this.m_dg[this.m_intCurrRow-1,2].ToString().Trim();
				p_uc.strVariant = this.m_dg[this.m_intCurrRow-1,1].ToString().Trim();
				p_uc.strCommonName = this.m_dg[this.m_intCurrRow-1,3].ToString().Trim();
				p_uc.strConvertedSpCd = this.m_dg[this.m_intCurrRow-1,4].ToString().Trim();
				p_uc.strComments = this.m_dg[this.m_intCurrRow-1,5].ToString().Trim();

			}
			System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
			if (result==System.Windows.Forms.DialogResult.OK)
			{
				if (p_strAction.Trim().ToUpper()=="NEW")
				{
					this.m_dv.AllowNew = true;
					System.Data.DataRow p_row =	this.m_ado.m_DataSet.Tables["tree_species_conversion"].NewRow();
					p_row["id"] = Convert.ToInt32(p_uc.strId);
					p_row["fvs_variant"] = p_uc.strVariant;
					p_row["fia_spcd"] = Convert.ToInt32(p_uc.strSpCd);
					p_row["convert_to_fia_spcd"] = Convert.ToInt32(p_uc.strConvertedSpCd);
					p_row["common_name"] = p_uc.strCommonName;
					p_row["comments"] = p_uc.strComments;
					this.m_ado.m_DataSet.Tables["tree_species_conversion"].Rows.Add(p_row);
					p_row=null;
					this.m_dv.AllowNew = false;

				}
				else
				{
					this.m_dg[this.m_intCurrRow-1,1] = p_uc.strVariant;
					this.m_dg[this.m_intCurrRow-1,2] = p_uc.strSpCd;
					this.m_dg[this.m_intCurrRow-1,3] = p_uc.strCommonName;
					this.m_dg[this.m_intCurrRow-1,4] = p_uc.strConvertedSpCd;
					this.m_dg[this.m_intCurrRow-1,5] = p_uc.strComments;

					this.m_dg.SetDataBinding(this.m_dv,"");
					this.m_dg.Update();

				}


				if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
				if (this.btnDelete.Enabled==false) this.btnDelete.Enabled=true;
				if (this.btnEdit.Enabled==false) this.btnEdit.Enabled=true;
			}
			frmTemp.Close();
			frmTemp.Dispose();
			frmTemp=null;

		}
		private int getUniqueId()
		{
            string strUniqueId="";
			int intId=0;
			int intId2=0;
			strUniqueId = this.m_ado.getSingleStringValueFromSQLQuery(this.m_ado.m_OleDbConnection,
				this.m_ado.m_OleDbTransaction,
				"select max(id) as maxid from " + this.m_strTreeSpcCvtTable,
				this.m_strTreeSpcCvtTable);

			if (strUniqueId != null && strUniqueId.Trim().Length > 0)
				intId = Convert.ToInt32(strUniqueId);

			if (this.m_ado.m_DataSet.Tables["tree_species_conversion"].Compute("Max(id)", null) != System.DBNull.Value)
			{
				intId2 = Convert.ToInt32(this.m_ado.m_DataSet.Tables["tree_species_conversion"].Compute("Max(id)", null));
			}
			if (intId2 >= intId) 
			{
				intId=intId2 + 1;
			}

			if (intId==0) return 1;
			return intId;

		}

		private void btnEdit_Click(object sender, System.EventArgs e)
		{
			this.EditForm("Edit");
		}
		private void val_data()
		{

			this.m_intError=0;
			string strCurVariant="";
			string strCurSpCd="";
			System.Data.DataTable p_dt = this.m_ado.m_DataSet.Tables["tree_species_conversion"];
			for (int x=0;x<=p_dt.Rows.Count-1;x++)
			{
				if (p_dt.Rows[x].RowState != System.Data.DataRowState.Deleted)
				{
					//all converted to spcd must have a value
					if (p_dt.Rows[x]["convert_to_fia_spcd"] == System.DBNull.Value ||
						p_dt.Rows[x]["convert_to_fia_spcd"].ToString().Trim().Length == 0)
					{
						MessageBox.Show("!!Missing value in column convert_to_fia_spcd!!","FIA Biosum",
							System.Windows.Forms.MessageBoxButtons.OK,
							System.Windows.Forms.MessageBoxIcon.Exclamation);

						this.m_intError=-1;
						return;
					}
					//make sure not converting softwood to a hardwood
					if (Convert.ToInt32(p_dt.Rows[x]["fia_spcd"]) < 300)
					{
					}

                    //make sure not converting hardwood to a softwood
					if (Convert.ToInt32(p_dt.Rows[x]["fia_spcd"]) > 299)
					{
					}

					strCurVariant=Convert.ToString(p_dt.Rows[x]["fvs_variant"]);
					strCurSpCd = Convert.ToString(p_dt.Rows[x]["fia_spcd"]);

					//make sure no duplicate variant + spcd combinations
					for (int y=x+1;y<=p_dt.Rows.Count-1;y++)
					{
						if (p_dt.Rows[y].RowState != System.Data.DataRowState.Deleted)
						{
							if (strCurVariant == Convert.ToString(p_dt.Rows[y]["fvs_variant"]))
							{
								if (strCurSpCd.Trim() == Convert.ToString(p_dt.Rows[y]["fia_spcd"]).Trim())
								{
									MessageBox.Show("!!Duplicate variant and tree species values found for variant " + strCurVariant + " and tree species " + strCurSpCd + ". Delete one of the records!!","FIA Biosum",
										System.Windows.Forms.MessageBoxButtons.OK,
										System.Windows.Forms.MessageBoxIcon.Exclamation);
									this.m_intError=-1;
									return;

								}
							}
						}
					}
				}
			}
		}


	}
	public class TreeSpcCvt_DataGridColoredTextBoxColumn : DataGridTextBoxColumn
	{
		bool m_bEdit=false;
		FIA_Biosum_Manager.uc_fvs_tree_spc_conversion uc_tree_spc_conversion1;
		string m_strLastKey="";
		bool m_bNumericOnly=false;
		

		public TreeSpcCvt_DataGridColoredTextBoxColumn(bool bEdit,bool bNumericOnly,FIA_Biosum_Manager.uc_fvs_tree_spc_conversion p_uc)
		{
			this.m_bEdit = bEdit;
			this.m_bNumericOnly = bNumericOnly;
			this.uc_tree_spc_conversion1 = p_uc;
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
							if (this.uc_tree_spc_conversion1.btnSave.Enabled==false) this.uc_tree_spc_conversion1.btnSave.Enabled=true;
						}
						else
						{
							if (e.KeyCode == Keys.Back)
							{
								this.m_strLastKey = Convert.ToString(e.KeyValue);
								if (this.uc_tree_spc_conversion1.btnSave.Enabled==false) this.uc_tree_spc_conversion1.btnSave.Enabled=true;
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
						if (this.uc_tree_spc_conversion1.btnSave.Enabled==false) this.uc_tree_spc_conversion1.btnSave.Enabled=true;
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
					if (this.uc_tree_spc_conversion1.btnSave.Enabled==false) this.uc_tree_spc_conversion1.btnSave.Enabled=true;
				}
			}
		}
		     
	}
}
