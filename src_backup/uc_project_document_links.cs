using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_project_document_links.
	/// </summary>
	public class uc_project_document_links : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.ToolBar tbrProjectDocumentLinks;
		private System.Windows.Forms.ToolBarButton btnGlobal;
		private System.Windows.Forms.ToolBarButton btnBoth;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.ContextMenu contextMenu1;
		private System.Windows.Forms.ToolBarButton btnView;
		private System.Windows.Forms.ToolBarButton btnPrivate;
		private System.Windows.Forms.MenuItem mnuViewAll;
		private System.Windows.Forms.MenuItem mnuViewCore;
		private System.Windows.Forms.MenuItem mnuViewFVS;
		private System.Windows.Forms.MenuItem mnuViewFRCS;
		private System.Windows.Forms.MenuItem mnuViewProcessor;
		private System.Windows.Forms.MenuItem mnuViewGIS;
		private System.Windows.Forms.StatusBar statusBar1;
		private System.Windows.Forms.StatusBarPanel sbQueryRecordCount;
		private System.Windows.Forms.StatusBarPanel sbDisplayedRecordCount;
		private System.Windows.Forms.ToolBarButton tbrSeparator1;
		private System.Windows.Forms.ToolBarButton tbrNew;
		private System.Windows.Forms.ToolBarButton tbrEdit;
		private System.Windows.Forms.ToolBarButton tbrDelete;
		private System.Windows.Forms.ToolBarButton tbrSeparator2;
		private System.Windows.Forms.ToolBarButton tbrOpen;
		private System.Windows.Forms.StatusBarPanel sbMsg;
		//private System.Windows.Forms.Button btnX;
		//private System.Windows.Forms.ToolBarButton btnX;
		public  string m_strProcess="";
		public string m_strSharedPrivate="";
		private System.Windows.Forms.MenuItem mnuViewOther;
		public string m_editmode="";
		public int m_intError=0;
		public string m_strError="";
		private System.Windows.Forms.ListView listView1;

		private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvRowColors = new ListViewAlternateBackgroundColors();

		public uc_project_document_links()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.listView1.Left = 2;
			this.tbrProjectDocumentLinks.Left = 4;
			this.listView1.Width = this.groupBox1.Width - 4;
			this.listView1.Height = this.statusBar1.Top - this.tbrProjectDocumentLinks.Top;
			this.contextMenu1.MenuItems[0].Checked=true;
			this.m_strProcess = "All";
			
			this.listView1.Clear();
			this.listView1.Columns.Add(" ",this.imageList1.Images[0].Width,HorizontalAlignment.Left);
			this.listView1.Columns.Add(" ", 30 , HorizontalAlignment.Center);
			this.listView1.Columns.Add("Document", 250, HorizontalAlignment.Left);
			this.listView1.Columns.Add("Description", 500, HorizontalAlignment.Left);
		
			this.sbMsg.Text="View all documents";
			this.sbQueryRecordCount.Text="";
			this.sbDisplayedRecordCount.Text="";
		    this.sbQueryRecordCount.Text ="";
			this.sbQueryRecordCount.Width  =
				(int)this.CreateGraphics().MeasureString("9999999/9999999",this.statusBar1.Font).Width;
		    this.sbMsg.Width = (int)(this.groupBox1.Width * .50) -  (int)(this.sbQueryRecordCount.Width * .50);
			this.sbDisplayedRecordCount.Width =(int)(this.groupBox1.Width * .50) - (int)(this.sbQueryRecordCount.Width * .50);
			this.sbDisplayedRecordCount.Alignment = System.Windows.Forms.HorizontalAlignment.Right;
			this.sbQueryRecordCount.Alignment = System.Windows.Forms.HorizontalAlignment.Center;

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

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(uc_project_document_links));
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.statusBar1 = new System.Windows.Forms.StatusBar();
			this.sbMsg = new System.Windows.Forms.StatusBarPanel();
			this.sbQueryRecordCount = new System.Windows.Forms.StatusBarPanel();
			this.sbDisplayedRecordCount = new System.Windows.Forms.StatusBarPanel();
			this.tbrProjectDocumentLinks = new System.Windows.Forms.ToolBar();
			this.btnGlobal = new System.Windows.Forms.ToolBarButton();
			this.btnPrivate = new System.Windows.Forms.ToolBarButton();
			this.btnBoth = new System.Windows.Forms.ToolBarButton();
			this.btnView = new System.Windows.Forms.ToolBarButton();
			this.contextMenu1 = new System.Windows.Forms.ContextMenu();
			this.mnuViewAll = new System.Windows.Forms.MenuItem();
			this.mnuViewCore = new System.Windows.Forms.MenuItem();
			this.mnuViewFVS = new System.Windows.Forms.MenuItem();
			this.mnuViewFRCS = new System.Windows.Forms.MenuItem();
			this.mnuViewProcessor = new System.Windows.Forms.MenuItem();
			this.mnuViewGIS = new System.Windows.Forms.MenuItem();
			this.mnuViewOther = new System.Windows.Forms.MenuItem();
			this.tbrSeparator1 = new System.Windows.Forms.ToolBarButton();
			this.tbrNew = new System.Windows.Forms.ToolBarButton();
			this.tbrEdit = new System.Windows.Forms.ToolBarButton();
			this.tbrDelete = new System.Windows.Forms.ToolBarButton();
			this.tbrSeparator2 = new System.Windows.Forms.ToolBarButton();
			this.tbrOpen = new System.Windows.Forms.ToolBarButton();
			this.lblTitle = new System.Windows.Forms.Label();
			this.listView1 = new System.Windows.Forms.ListView();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.sbMsg)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.sbQueryRecordCount)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.sbDisplayedRecordCount)).BeginInit();
			this.SuspendLayout();
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.listView1);
			this.groupBox1.Controls.Add(this.statusBar1);
			this.groupBox1.Controls.Add(this.tbrProjectDocumentLinks);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(696, 424);
			this.groupBox1.TabIndex = 30;
			this.groupBox1.TabStop = false;
			this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
			// 
			// statusBar1
			// 
			this.statusBar1.Location = new System.Drawing.Point(3, 397);
			this.statusBar1.Name = "statusBar1";
			this.statusBar1.Panels.AddRange(new System.Windows.Forms.StatusBarPanel[] {
																						  this.sbMsg,
																						  this.sbQueryRecordCount,
																						  this.sbDisplayedRecordCount});
			this.statusBar1.ShowPanels = true;
			this.statusBar1.Size = new System.Drawing.Size(690, 24);
			this.statusBar1.SizingGrip = false;
			this.statusBar1.TabIndex = 34;
			// 
			// sbMsg
			// 
			this.sbMsg.Text = "sbMsg";
			// 
			// sbQueryRecordCount
			// 
			this.sbQueryRecordCount.Text = "sbQueryRecordCount";
			// 
			// sbDisplayedRecordCount
			// 
			this.sbDisplayedRecordCount.Text = "sbDisplayedRecordCount";
			// 
			// tbrProjectDocumentLinks
			// 
			this.tbrProjectDocumentLinks.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.tbrProjectDocumentLinks.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
																									   this.btnGlobal,
																									   this.btnPrivate,
																									   this.btnBoth,
																									   this.btnView,
																									   this.tbrSeparator1,
																									   this.tbrNew,
																									   this.tbrEdit,
																									   this.tbrDelete,
																									   this.tbrSeparator2,
																									   this.tbrOpen});
			this.tbrProjectDocumentLinks.ButtonSize = new System.Drawing.Size(50, 36);
			this.tbrProjectDocumentLinks.Divider = false;
			this.tbrProjectDocumentLinks.Dock = System.Windows.Forms.DockStyle.None;
			this.tbrProjectDocumentLinks.DropDownArrows = true;
			this.tbrProjectDocumentLinks.ImageList = this.imageList1;
			this.tbrProjectDocumentLinks.Location = new System.Drawing.Point(8, 48);
			this.tbrProjectDocumentLinks.Name = "tbrProjectDocumentLinks";
			this.tbrProjectDocumentLinks.ShowToolTips = true;
			this.tbrProjectDocumentLinks.Size = new System.Drawing.Size(680, 41);
			this.tbrProjectDocumentLinks.TabIndex = 31;
			this.tbrProjectDocumentLinks.Click += new System.EventHandler(this.tbrProjectDocumentLinks_Click);
			this.tbrProjectDocumentLinks.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.tbrProjectDocumentLinks_ButtonClick);
			// 
			// btnGlobal
			// 
			this.btnGlobal.ImageIndex = 0;
			this.btnGlobal.Text = "Shared";
			// 
			// btnPrivate
			// 
			this.btnPrivate.ImageIndex = 1;
			this.btnPrivate.Text = "Private";
			// 
			// btnBoth
			// 
			this.btnBoth.ImageIndex = 2;
			this.btnBoth.Pushed = true;
			this.btnBoth.Text = "Both";
			// 
			// btnView
			// 
			this.btnView.DropDownMenu = this.contextMenu1;
			this.btnView.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton;
			this.btnView.Text = "View";
			// 
			// contextMenu1
			// 
			this.contextMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																						 this.mnuViewAll,
																						 this.mnuViewCore,
																						 this.mnuViewFVS,
																						 this.mnuViewFRCS,
																						 this.mnuViewProcessor,
																						 this.mnuViewGIS,
																						 this.mnuViewOther});
			// 
			// mnuViewAll
			// 
			this.mnuViewAll.Index = 0;
			this.mnuViewAll.Text = "All";
			this.mnuViewAll.Click += new System.EventHandler(this.mnuViewAll_Click);
			// 
			// mnuViewCore
			// 
			this.mnuViewCore.Index = 1;
			this.mnuViewCore.Text = "Core Analysis";
			this.mnuViewCore.Click += new System.EventHandler(this.mnuViewCore_Click);
			// 
			// mnuViewFVS
			// 
			this.mnuViewFVS.Index = 2;
			this.mnuViewFVS.Text = "FVS";
			this.mnuViewFVS.Click += new System.EventHandler(this.mnuViewFVS_Click);
			// 
			// mnuViewFRCS
			// 
			this.mnuViewFRCS.Index = 3;
			this.mnuViewFRCS.Text = "FRCS";
			this.mnuViewFRCS.Click += new System.EventHandler(this.mnuViewFRCS_Click);
			// 
			// mnuViewProcessor
			// 
			this.mnuViewProcessor.Index = 4;
			this.mnuViewProcessor.Text = "Processor";
			this.mnuViewProcessor.Click += new System.EventHandler(this.mnuViewProcessor_Click);
			// 
			// mnuViewGIS
			// 
			this.mnuViewGIS.Index = 5;
			this.mnuViewGIS.Text = "GIS";
			this.mnuViewGIS.Click += new System.EventHandler(this.mnuViewGIS_Click);
			// 
			// mnuViewOther
			// 
			this.mnuViewOther.Index = 6;
			this.mnuViewOther.Text = "Other";
			this.mnuViewOther.Click += new System.EventHandler(this.mnuViewOther_Click);
			// 
			// tbrSeparator1
			// 
			this.tbrSeparator1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
			// 
			// tbrNew
			// 
			this.tbrNew.ImageIndex = 3;
			this.tbrNew.Text = "New";
			// 
			// tbrEdit
			// 
			this.tbrEdit.ImageIndex = 4;
			this.tbrEdit.Text = "Edit";
			// 
			// tbrDelete
			// 
			this.tbrDelete.ImageIndex = 6;
			this.tbrDelete.Text = "Delete";
			// 
			// tbrSeparator2
			// 
			this.tbrSeparator2.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
			// 
			// tbrOpen
			// 
			this.tbrOpen.ImageIndex = 5;
			this.tbrOpen.Text = "Open";
			// 
			// lblTitle
			// 
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(8, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(352, 24);
			this.lblTitle.TabIndex = 30;
			this.lblTitle.Text = "Project Document Links Depository";
			// 
			// listView1
			// 
			this.listView1.GridLines = true;
			this.listView1.Location = new System.Drawing.Point(8, 112);
			this.listView1.Name = "listView1";
			this.listView1.Size = new System.Drawing.Size(680, 280);
			this.listView1.SmallImageList = this.imageList1;
			this.listView1.TabIndex = 35;
			this.listView1.View = System.Windows.Forms.View.Details;
			this.listView1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.listView1_MouseUp);
			this.listView1.SelectedIndexChanged += new System.EventHandler(this.listView1_SelectedIndexChanged);
			// 
			// uc_project_document_links
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_project_document_links";
			this.Size = new System.Drawing.Size(696, 424);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.sbMsg)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.sbQueryRecordCount)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.sbDisplayedRecordCount)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private void tbrProjectDocumentLinks_Click(object sender, System.EventArgs e)
		{
			
		}
		private void setCurrentViewText()
		{
			const byte SHARED = 0;
            const byte PRIVATE = 1;
			//const byte BOTH = 2;
			const byte ALL = 0;
			const byte CORE = 1;
			const byte FVS = 2;
			const byte FRCS = 3;
			const byte PROCESSOR = 4;
			const byte GIS = 5;
			string str1="";
			string str2="";
			int intSelection;

			if (this.tbrProjectDocumentLinks.Buttons[SHARED].Pushed==true)
			{
                str2 ="shared";
			}
			else if (this.tbrProjectDocumentLinks.Buttons[PRIVATE].Pushed==true)
			{
                str2 = "private";
			}
			else
			{
				str2 = "shared and private";
			}

			if (this.contextMenu1.MenuItems[ALL].Checked==true)
			{
				str1 = "all";
				intSelection = ALL;
				
			}
			else if (this.contextMenu1.MenuItems[CORE].Checked==true)
			{
				str1 = "core analysis";
				intSelection = CORE;
			}
			else if (this.contextMenu1.MenuItems[FVS].Checked==true)
			{
				str1 = "fvs";
				intSelection = FVS;
			}
			else if (this.contextMenu1.MenuItems[FRCS].Checked==true)
			{
				str1= "frcs";
				intSelection = FRCS;
			}
			else if (this.contextMenu1.MenuItems[PROCESSOR].Checked==true)
			{
				str1="processor";
				intSelection = PROCESSOR;
			}
			else if (this.contextMenu1.MenuItems[GIS].Checked==true)
			{
				str1="gis";
				intSelection = GIS;
			}
			else
			{
				str1="miscellaneous";
				intSelection = 6;
			}

			this.sbMsg.Text = "View " + str1 + " " + str2 + " documents";
            this.m_strProcess = this.contextMenu1.MenuItems[intSelection].Text;

		}
		private void tbrProjectDocumentLinks_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			DialogResult result = new DialogResult();
			frmDialog frmtemp = new frmDialog();
			try
			{
				switch (e.Button.Text.Trim().ToUpper())
				{
					case "SHARED":
						this.tbrProjectDocumentLinks.Buttons[0].Pushed=true;	    
						this.tbrProjectDocumentLinks.Buttons[1].Pushed=false;	    
						this.tbrProjectDocumentLinks.Buttons[2].Pushed=false;	
						this.loadvalues(((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text,
							((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text,
							true);
						break;
					case "PRIVATE":
						this.tbrProjectDocumentLinks.Buttons[0].Pushed=false;	    
						this.tbrProjectDocumentLinks.Buttons[1].Pushed=true;	    
						this.tbrProjectDocumentLinks.Buttons[2].Pushed=false;	 
						this.loadvalues(((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text,
							((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text,
							true);

						break;
					case "BOTH":
						this.tbrProjectDocumentLinks.Buttons[0].Pushed=false;	    
						this.tbrProjectDocumentLinks.Buttons[1].Pushed=false;	    
						this.tbrProjectDocumentLinks.Buttons[2].Pushed=true;	
						this.loadvalues(((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text,
							((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text,
							true);

						break;
					case "NEW":
						
						frmtemp.Text = "New Document Link";
						//frmtemp.uc_project1.Visible=false;
						//frmtemp.uc_scenario1.Visible=false;
						//frmtemp.uc_select_list_item1.Visible=false;
						//frmtemp.uc_project_document_links1.Visible=false;
						frmtemp.uc_project_document_links_edit1.Visible=true;
						frmtemp.uc_project_document_links_edit1.cmbCategory.Text = this.m_strProcess;
						if (((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text.Trim().Length > 0) 
						{
							if (!System.IO.Directory.Exists(
								((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text.Trim()))
								frmtemp.uc_project_document_links_edit1.chkShared.Enabled=false;
						}
						else frmtemp.uc_project_document_links_edit1.chkShared.Enabled=false;
						if (((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text.Trim().Length > 0) 
						{
							if (!System.IO.Directory.Exists(
								((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text.Trim()))
								frmtemp.uc_project_document_links_edit1.chkPrivate.Enabled=false;
						}
						else frmtemp.uc_project_document_links_edit1.chkPrivate.Enabled=false;
						frmtemp.uc_project_document_links_edit1.strOKButtonText = "OK";
						frmtemp.uc_project_document_links_edit1.strCancelButtonText="Cancel";
						result = frmtemp.ShowDialog(this);
					switch (result)
					{
						case DialogResult.OK:

							this.addToList(
								frmtemp.uc_project_document_links_edit1.txtDocument.Text,
								frmtemp.uc_project_document_links_edit1.txtDescription.Text,
								frmtemp.uc_project_document_links_edit1.cmbCategory.Text,
								frmtemp.uc_project_document_links_edit1.m_strSharedPrivate);
                                
							this.savevalues(
								"n",
								frmtemp.uc_project_document_links_edit1.txtDocument.Text,
								frmtemp.uc_project_document_links_edit1.txtDescription.Text,
								frmtemp.uc_project_document_links_edit1.cmbCategory.Text,
								frmtemp.uc_project_document_links_edit1.m_strSharedPrivate);
								

							break;
						default:
							break;
					}
						frmtemp.Close();
						frmtemp = null;

						break;
					case "EDIT":
						//MessageBox.Show(this.listView1.SelectedItems[0].ImageIndex.ToString());
						
						
						frmtemp.Text = "Edit Document Link";
						//frmtemp.uc_project1.Visible=false;
						//frmtemp.uc_scenario1.Visible=false;
						//frmtemp.uc_select_list_item1.Visible=false;
						//frmtemp.uc_project_document_links1.Visible=false;
						frmtemp.uc_project_document_links_edit1.Visible=true;
						frmtemp.uc_project_document_links_edit1.cmbCategory.Text = this.m_strProcess;
						
						if (this.listView1.SelectedItems[0].ImageIndex == 0)
						{
							frmtemp.uc_project_document_links_edit1.chkPrivate.Checked=false;
							frmtemp.uc_project_document_links_edit1.chkPrivate.Enabled=false;
							frmtemp.uc_project_document_links_edit1.chkShared.Checked=true;
							frmtemp.uc_project_document_links_edit1.chkShared.Enabled=false;
						}
						else 
						{
							frmtemp.uc_project_document_links_edit1.chkPrivate.Checked=true;
							frmtemp.uc_project_document_links_edit1.chkPrivate.Enabled=false;
							frmtemp.uc_project_document_links_edit1.chkShared.Checked=false;
							frmtemp.uc_project_document_links_edit1.chkShared.Enabled=false;
						}
						frmtemp.uc_project_document_links_edit1.txtDescription.Text = 
							this.listView1.SelectedItems[0].SubItems[3].Text;
						frmtemp.uc_project_document_links_edit1.txtDocument.Text = 
							this.listView1.SelectedItems[0].SubItems[2].Text;
					switch (this.listView1.SelectedItems[0].SubItems[1].Text)
					{
						case "MISC":
							frmtemp.uc_project_document_links_edit1.cmbCategory.Text = "Other";
							break;
						case "CORE":
							frmtemp.uc_project_document_links_edit1.cmbCategory.Text = "Core Analysis";
							break;
						case "FVS":
							frmtemp.uc_project_document_links_edit1.cmbCategory.Text = "FVS";
							break;
						case "FRCS":
							frmtemp.uc_project_document_links_edit1.cmbCategory.Text = "FRCS";
							break;
						case "PROCESSOR":
							frmtemp.uc_project_document_links_edit1.cmbCategory.Text = "Processor";
							break;
						case "GIS":
							frmtemp.uc_project_document_links_edit1.cmbCategory.Text = "GIS";
							break;
						default:
							frmtemp.uc_project_document_links_edit1.cmbCategory.Text = "Core Analysis";
							break;
					}
					frmtemp.uc_project_document_links_edit1.strOKButtonText = "OK";
					frmtemp.uc_project_document_links_edit1.strCancelButtonText="Cancel";
					result = frmtemp.ShowDialog(this);
					switch (result)
					{
						case DialogResult.OK:
							//see if there were any changes
							if ((frmtemp.uc_project_document_links_edit1.txtDocument.Text.Trim().ToUpper()
								!=
								this.listView1.SelectedItems[0].SubItems[2].Text.Trim().ToUpper())
								||
								(frmtemp.uc_project_document_links_edit1.cmbCategory.Text.Trim().ToUpper()
								!=
								this.getListToComboMenuText(this.listView1.SelectedItems[0].SubItems[1].Text.Trim().ToUpper()))
								||
								(frmtemp.uc_project_document_links_edit1.txtDescription.Text.Trim().ToUpper()
								!=
								this.listView1.SelectedItems[0].SubItems[3].Text.Trim().ToUpper()))
							{
								//MessageBox.Show("made it");

								this.savevalues(
									"e",
									frmtemp.uc_project_document_links_edit1.txtDocument.Text,
									frmtemp.uc_project_document_links_edit1.txtDescription.Text,
									frmtemp.uc_project_document_links_edit1.cmbCategory.Text,
									frmtemp.uc_project_document_links_edit1.m_strSharedPrivate);

								if (this.m_intError==0)
								{
									if (this.getViewMenuText(this.listView1.SelectedItems[0].SubItems[1].Text).Trim().ToUpper()
										==
										this.getViewMenuText(frmtemp.uc_project_document_links_edit1.cmbCategory.Text.Trim().ToUpper()) 
										|| this.RemoveFromCurrentList(frmtemp.uc_project_document_links_edit1.cmbCategory.Text)==false)
									{

										this.listView1.SelectedItems[0].SubItems[3].Text =
											frmtemp.uc_project_document_links_edit1.txtDescription.Text;
										this.listView1.SelectedItems[0].SubItems[2].Text = 
											frmtemp.uc_project_document_links_edit1.txtDocument.Text;
										this.listView1.SelectedItems[0].SubItems[1].Text = 
											this.getViewMenuText(frmtemp.uc_project_document_links_edit1.cmbCategory.Text);
									}
									else
									{
										this.loadvalues(((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text,
											((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text,
											true);
									}
								}
								

							}

							
                                

							break;
						default:
							break;
					}
						frmtemp.Close();
						frmtemp = null;

                     

						break;
					
					case "DELETE":

						frmtemp.Text = "Delete Document Link";
						//frmtemp.uc_project1.Visible=false;
						//frmtemp.uc_scenario1.Visible=false;
						//frmtemp.uc_select_list_item1.Visible=false;
						//frmtemp.uc_project_document_links1.Visible=false;
						frmtemp.uc_project_document_links_edit1.Visible=true;
						frmtemp.uc_project_document_links_edit1.cmbCategory.Text = this.m_strProcess;
						
						if (this.listView1.SelectedItems[0].ImageIndex == 0)
						{
							frmtemp.uc_project_document_links_edit1.chkPrivate.Checked=false;
							frmtemp.uc_project_document_links_edit1.chkPrivate.Enabled=false;
							frmtemp.uc_project_document_links_edit1.chkShared.Checked=true;
							frmtemp.uc_project_document_links_edit1.chkShared.Enabled=false;
						}
						else 
						{
							frmtemp.uc_project_document_links_edit1.chkPrivate.Checked=true;
							frmtemp.uc_project_document_links_edit1.chkPrivate.Enabled=false;
							frmtemp.uc_project_document_links_edit1.chkShared.Checked=false;
							frmtemp.uc_project_document_links_edit1.chkShared.Enabled=false;
						}
						frmtemp.uc_project_document_links_edit1.txtDescription.Text = 
							this.listView1.SelectedItems[0].SubItems[3].Text;
						frmtemp.uc_project_document_links_edit1.txtDocument.Text = 
							this.listView1.SelectedItems[0].SubItems[2].Text;
					    switch (this.listView1.SelectedItems[0].SubItems[1].Text)
					    {
						   case "MISC":
						    	frmtemp.uc_project_document_links_edit1.cmbCategory.Text = "Other";
							    break;
						   case "CORE":
						    	frmtemp.uc_project_document_links_edit1.cmbCategory.Text = "Core Analysis";
						    	break;
						   case "FVS":
						    	frmtemp.uc_project_document_links_edit1.cmbCategory.Text = "FVS";
							    break;
						   case "FRCS":
						    	frmtemp.uc_project_document_links_edit1.cmbCategory.Text = "FRCS";
							    break;
						   case "PROCESSOR":
						    	frmtemp.uc_project_document_links_edit1.cmbCategory.Text = "Processor";
						    	break;
						   case "GIS":
						    	frmtemp.uc_project_document_links_edit1.cmbCategory.Text = "GIS";
							    break;
						   default:
						    	frmtemp.uc_project_document_links_edit1.cmbCategory.Text = "Core Analysis";
							    break;
					    }
						frmtemp.uc_project_document_links_edit1.txtDescription.Enabled=false;
						frmtemp.uc_project_document_links_edit1.txtDocument.Enabled=false;
						frmtemp.uc_project_document_links_edit1.cmbCategory.Enabled=false;

					    frmtemp.uc_project_document_links_edit1.lblMsg.Font= new Font(this.Font, FontStyle.Bold);
						frmtemp.uc_project_document_links_edit1.lblMsg.ForeColor = Color.Red;
						frmtemp.uc_project_document_links_edit1.lblMsg.Text = "Delete this record? (Y/N)";
						frmtemp.uc_project_document_links_edit1.strOKButtonText = "Yes";
						frmtemp.uc_project_document_links_edit1.strCancelButtonText="No";

						result = frmtemp.ShowDialog(this);
					   switch (result)
					   {
						  case DialogResult.OK:
							  this.savevalues(
								  "d",
								  frmtemp.uc_project_document_links_edit1.txtDocument.Text,
								  frmtemp.uc_project_document_links_edit1.txtDescription.Text,
								  frmtemp.uc_project_document_links_edit1.cmbCategory.Text,
								  frmtemp.uc_project_document_links_edit1.m_strSharedPrivate);
							  this.loadvalues(((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text,
								  ((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text,
								  true);
							   break;
					   }

					   frmtemp.Close();
					   frmtemp = null;
						break;
				
					case "OPEN":

						System.Diagnostics.Process.Start(
							this.listView1.SelectedItems[0].SubItems[2].Text);
						break;
                 

				}
			}
			catch (Exception caught)
			{
				MessageBox.Show(caught.Message,"FIA Biosum Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
			}
			this.setCurrentViewText();


		}

		private void mnuViewAll_Click(object sender, System.EventArgs e)
		{

			this.mnuViewAll.Checked=true;
			this.mnuViewCore.Checked=false;
			this.mnuViewFVS.Checked=false;
			this.mnuViewGIS.Checked=false;
			this.mnuViewProcessor.Checked=false;
			this.mnuViewFRCS.Checked=false;			
			this.mnuViewOther.Checked=false;
			this.setCurrentViewText();
			this.loadvalues(((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text,
				((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text,
				true);
		}

		private void mnuViewCore_Click(object sender, System.EventArgs e)
		{
			this.mnuViewAll.Checked=false;
			this.mnuViewCore.Checked=true;
			this.mnuViewFVS.Checked=false;
			this.mnuViewGIS.Checked=false;
			this.mnuViewProcessor.Checked=false;
			this.mnuViewFRCS.Checked=false;
			this.mnuViewOther.Checked=false;
			this.setCurrentViewText();
			this.loadvalues(((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text,
				((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text,
				true);

		}

		private void mnuViewFVS_Click(object sender, System.EventArgs e)
		{
			this.mnuViewAll.Checked=false;
			this.mnuViewCore.Checked=false;
			this.mnuViewFVS.Checked=true;
			this.mnuViewGIS.Checked=false;
			this.mnuViewProcessor.Checked=false;
			this.mnuViewFRCS.Checked=false;
			this.mnuViewOther.Checked=false;
			this.setCurrentViewText();
			this.loadvalues(((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text,
				((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text,
				true);

		}

		private void mnuViewFRCS_Click(object sender, System.EventArgs e)
		{
			this.mnuViewAll.Checked=false;
			this.mnuViewCore.Checked=false;
			this.mnuViewFVS.Checked=false;
			this.mnuViewGIS.Checked=false;
			this.mnuViewProcessor.Checked=false;
			this.mnuViewFRCS.Checked=true;
			this.mnuViewOther.Checked=false;
			this.setCurrentViewText();
			this.loadvalues(((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text,
				((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text,
				true);

		}

		private void mnuViewProcessor_Click(object sender, System.EventArgs e)
		{
			this.mnuViewAll.Checked=false;
			this.mnuViewCore.Checked=false;
			this.mnuViewFVS.Checked=false;
			this.mnuViewGIS.Checked=false;
			this.mnuViewProcessor.Checked=true;
			this.mnuViewFRCS.Checked=false;
			this.mnuViewOther.Checked=false;
            this.setCurrentViewText();
			this.loadvalues(((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text,
				((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text,
				true);
		}

		private void mnuViewGIS_Click(object sender, System.EventArgs e)
		{
			this.mnuViewAll.Checked=false;
			this.mnuViewCore.Checked=false;
			this.mnuViewFVS.Checked=false;
			this.mnuViewGIS.Checked=true;
			this.mnuViewProcessor.Checked=false;
			this.mnuViewFRCS.Checked=false;
			this.mnuViewOther.Checked=false;
			this.setCurrentViewText();
			this.loadvalues(((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text,
				((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text,
				true);

		}

		private void mnuViewOther_Click(object sender, System.EventArgs e)
		{
			this.mnuViewAll.Checked=false;
			this.mnuViewCore.Checked=false;
			this.mnuViewFVS.Checked=false;
			this.mnuViewGIS.Checked=false;
			this.mnuViewProcessor.Checked=false;
			this.mnuViewFRCS.Checked=false;
			this.mnuViewOther.Checked=true;
			this.setCurrentViewText();
			this.loadvalues(((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text,
				((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text,
				true);
		}

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.listView1.Width = this.groupBox1.Width - 4;
				this.tbrProjectDocumentLinks.Width = this.groupBox1.Width - 8;
				this.sbQueryRecordCount.Width  =
					(int)this.CreateGraphics().MeasureString("9999999/9999999",this.statusBar1.Font).Width;
				this.sbMsg.Width = (int)(this.groupBox1.Width * .50) -  (int)(this.sbQueryRecordCount.Width * .50);
				this.sbDisplayedRecordCount.Width =(int)(this.groupBox1.Width * .50) - (int)(this.sbQueryRecordCount.Width * .50);
				this.listView1.Height =   this.statusBar1.Top - this.listView1.Top ;//this.tbrProjectDocumentLinks.Top - this.tbrProjectDocumentLinks.Height ;
			}
			catch 
			{
			}
		}
		private void loadsharedlinks(string strDir, string strSQLWhere)
		{
			string strMDB=strDir.Trim() + "\\shared_project_links_and_notes.mdb";
			string strSQL="";
			string strConn="";
			int x=0;


		    ado_data_access p_ado = new ado_data_access();
			

			strConn=p_ado.getMDBConnString(strMDB,"admin","");
			//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDB + ";User Id=admin;Password=;";
			p_ado.OpenConnection(strConn);	
			if (p_ado.m_intError != 0)
			{

				p_ado = null;
				return;
			}
			if (strSQLWhere.Trim().Length > 0)
			{
				strSQL = "SELECT * FROM links_depository WHERE " + 
					 strSQLWhere + " AND list_YN ='Y';";
			}
			else 
			{
				strSQL = "SELECT * FROM links_depository WHERE list_YN = 'Y';";
			}
			p_ado.SqlQueryReader(p_ado.m_OleDbConnection,strSQL);

			if (p_ado.m_intError != 0)
			{

				p_ado.m_OleDbConnection.Close();
				p_ado = null;
				return;
			}
			if (listView1.Items.Count > 0)
				x=listView1.Items.Count-1;
			else x=0;
			while (p_ado.m_OleDbDataReader.Read())
			{
				if (p_ado.m_OleDbDataReader["link"] != System.DBNull.Value)
				{
					if (p_ado.m_OleDbDataReader["link"].ToString().Trim().Length > 0)
					{
						System.Windows.Forms.ListViewItem entryListItem =
							listView1.Items.Add(" ",0);


						this.m_oLvRowColors.AddRow();
						this.m_oLvRowColors.AddColumns(x,listView1.Columns.Count);

						entryListItem.UseItemStyleForSubItems=false;
						this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,0,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
                        
						listView1.Items[x].SubItems.Add(this.getViewMenuText(p_ado.m_OleDbDataReader["category"].ToString().Trim()));
						m_oLvRowColors.ListViewSubItem(entryListItem.Index,1,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
						listView1.Items[x].SubItems.Add(p_ado.m_OleDbDataReader["link"].ToString().Trim());
						m_oLvRowColors.ListViewSubItem(entryListItem.Index,2,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
						listView1.Items[x].SubItems.Add(p_ado.m_OleDbDataReader["description"].ToString().Trim());
						m_oLvRowColors.ListViewSubItem(entryListItem.Index,3,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
						listView1.Items[x].SubItems[0] = entryListItem.SubItems[0];
						x++;

					}
				}
					
			}
			p_ado.m_OleDbDataReader.Close();
			p_ado.m_OleDbConnection.Close();
			p_ado = null;


		}
		private void loadprivatelinks(string strDir, string strSQLWhere)
		{
			string strMDB=strDir.Trim() + "\\personal_project_links_and_notes.mdb";
			string strSQL="";
			string strConn="";
			int x=0;
			//string str="";
			ado_data_access p_ado = new ado_data_access();
			

			strConn=p_ado.getMDBConnString(strMDB,"admin","");
			//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDB + ";User Id=admin;Password=;";
			p_ado.OpenConnection(strConn);	
			if (p_ado.m_intError != 0)
			{

				p_ado = null;
				return;
			}
			if (strSQLWhere.Trim().Length > 0)
			{
				strSQL = "SELECT * FROM links_depository WHERE " + 
					 strSQLWhere + " AND list_yn='Y';";
			}
			else 
			{
				strSQL = "SELECT * FROM links_depository WHERE list_yn='Y';";
			}
			p_ado.SqlQueryReader(p_ado.m_OleDbConnection,strSQL);

			if (p_ado.m_intError != 0)
			{

				p_ado.m_OleDbConnection.Close();
				p_ado = null;
				return;
			}

			if (listView1.Items.Count > 0)
				x=listView1.Items.Count;
			else x=0;

			while (p_ado.m_OleDbDataReader.Read())
			{
				if (p_ado.m_OleDbDataReader["link"] != System.DBNull.Value)
				{
					if (p_ado.m_OleDbDataReader["link"].ToString().Trim().Length > 0)
					{
						System.Windows.Forms.ListViewItem entryListItem =
							listView1.Items.Add(" ",1);


						this.m_oLvRowColors.AddRow();
						this.m_oLvRowColors.AddColumns(x,listView1.Columns.Count);

						entryListItem.UseItemStyleForSubItems=false;
						this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,0,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
                        
						listView1.Items[x].SubItems.Add(this.getViewMenuText(p_ado.m_OleDbDataReader["category"].ToString().Trim()));
						m_oLvRowColors.ListViewSubItem(entryListItem.Index,1,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
						listView1.Items[x].SubItems.Add(p_ado.m_OleDbDataReader["link"].ToString().Trim());
						m_oLvRowColors.ListViewSubItem(entryListItem.Index,2,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
						listView1.Items[x].SubItems.Add(p_ado.m_OleDbDataReader["description"].ToString().Trim());
						m_oLvRowColors.ListViewSubItem(entryListItem.Index,3,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
						listView1.Items[x].SubItems[0] = entryListItem.SubItems[0];
						x++;

					}
				}
					
			}
			p_ado.m_OleDbDataReader.Close();
			p_ado.m_OleDbConnection.Close();
			p_ado = null;


		}
		private string getViewMenuSQL()
		{
			string strSQL="";
			if (this.mnuViewAll.Checked==true)
			{
			}
			else if (this.mnuViewFVS.Checked==true)
			{
               strSQL = " category = 1";				
			
			}
			else if (this.mnuViewGIS.Checked==true)
			{
			   strSQL = " category = 2";
			}
			else if (this.mnuViewFRCS.Checked==true)
			{
			   strSQL = " category = 4";
			}
			else if (this.mnuViewOther.Checked==true)
			{
			   strSQL = " category = 9";
			}
			else if (this.mnuViewCore.Checked==true)
			{
			   strSQL = " category = 3";
			}
			else if (this.mnuViewProcessor.Checked==true)
			{
			   strSQL = " category = 5";
			}
            //if (strSQL.Trim().Length > 0) strSQL = " WHERE " + strSQL;
			return strSQL;
		}
		private string getViewMenuText(string strCategory)
		{
			switch (strCategory.Trim().ToUpper())
			{
				case "1":
					return "FVS";
				case "2":
					return "GIS";
				case "3":
					return "CORE";
				case "4":
					return "FRCS";
				case "5":
					return "PROCESSOR";
				case "9":
					return "MISC";
				case "CORE ANALYSIS":
					return "CORE";
				case "FVS":
					return "FVS";
				case "FRCS":
					return "FRCS";
				case "PROCESSOR":
					return "PROCESSOR";
                case "GIS":
					return "GIS";
				case "OTHER":
					return "MISC";
			}
			return "UNK";

		}

		private string getListToComboMenuText(string strCategory)
		{
			switch (strCategory.Trim().ToUpper())
			{
				case "CORE":
					return "CORE ANALYSIS";
				case "FVS":
					return "FVS";
				case "FRCS":
					return "FRCS";
				case "PROCESSOR":
					return "PROCESSOR";
				case "GIS":
					return "GIS";
				case "OTHER":
					return "OTHER";
			}
			return "OTHER";

		}

		private byte getCategoryAssignment(string strCategory)
		{

			const byte CORE = 3;
			const byte FVS = 1;
			const byte FRCS = 4;
			const byte PROCESSOR = 5;
			const byte GIS = 2;
			const byte OTHER = 9;

			//int x=0;

			switch (strCategory.Trim().ToUpper())
			{
				case "CORE ANALYSIS":
					return CORE;
                case "CORE":
					return CORE;
				case "FVS":
					return  FVS;
				case "FRCS":
					return FRCS;
				case "PROCESSOR":
					return PROCESSOR;
				case "GIS":
					return GIS;
                case "MISC":
					 return OTHER;
				case "OTHER":
					return OTHER;
			}
			return OTHER;

		}

		private int getSubCategoryAssignment(string strCategory)
		{

			const int CORE = 200;
			const int FVS = 1;
			const int FRCS = 300;
			const int PROCESSOR = 400;
			const int GIS = 100;
			const int OTHER = 999;

			//int x=0;

			switch (strCategory.Trim().ToUpper())
			{
				case "CORE ANALYSIS":
					return CORE;
                case "CORE":
					return CORE;
				case "FVS":
					return  FVS;
				case "FRCS":
					return FRCS;
				case "PROCESSOR":
					return PROCESSOR;
				case "GIS":
					return GIS;
                case "MISC":
					return OTHER;
				case "OTHER":
					return OTHER;
			}
			return OTHER;

		}

		public void addToList(string strDocument, string strDesc,string strCategory, string strSharedPrivate)
		{
			const byte ALL = 0;
			//const byte CORE = 1;
			//const byte FVS = 2;
			//const byte FRCS = 3;
			//const byte PROCESSOR = 4;
			//const byte GIS = 5;
			//const byte OTHER = 6;
            int x=0;
            bool lList=false;
               
			
			//see if the current document can be listed based on the current view

			if (this.contextMenu1.MenuItems[ALL].Checked==true) lList=true;
			else
			{
				for (x=0; x<= this.contextMenu1.MenuItems.Count-1;x++) 
				{
					if (this.contextMenu1.MenuItems[x].Checked == true &&
						this.contextMenu1.MenuItems[x].Text.Trim().ToUpper() == 
						strCategory.Trim().ToUpper()) 
					{
						lList=true;
						break;
					}
				}
			}
			System.Windows.Forms.ListView entryListItem=null;

			if (lList==true)
			{
				switch (strSharedPrivate)
				{
					case "S":
						
						this.listView1.Items.Add("",0);
						break;
					case "P":
						this.listView1.Items.Add("",1);
						break;
					default:   //B = both
						this.listView1.Items.Add("",0);
						break;
				}
				this.m_oLvRowColors.AddRow();
				this.m_oLvRowColors.AddColumns(listView1.Items.Count-1,listView1.Columns.Count);

				listView1.Items[listView1.Items.Count-1].UseItemStyleForSubItems=false;

				this.m_oLvRowColors.ListViewSubItem(listView1.Items[listView1.Items.Count-1].Index,0,listView1.Items[listView1.Items.Count-1].SubItems[listView1.Items[listView1.Items.Count-1].SubItems.Count-1],false);

				
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(this.getViewMenuText(strCategory.Trim().ToUpper()));
				this.m_oLvRowColors.ListViewSubItem(listView1.Items[listView1.Items.Count-1].Index,1,listView1.Items[listView1.Items.Count-1].SubItems[listView1.Items[listView1.Items.Count-1].SubItems.Count-1],false);
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(strDocument);
				this.m_oLvRowColors.ListViewSubItem(listView1.Items[listView1.Items.Count-1].Index,2,listView1.Items[listView1.Items.Count-1].SubItems[listView1.Items[listView1.Items.Count-1].SubItems.Count-1],false);
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(strDesc);		
				this.m_oLvRowColors.ListViewSubItem(listView1.Items[listView1.Items.Count-1].Index,3,listView1.Items[listView1.Items.Count-1].SubItems[listView1.Items[listView1.Items.Count-1].SubItems.Count-1],false);
                
				//two entries for both shared and private
				if (strSharedPrivate == "B")
				{
					this.listView1.Items.Add("",1);


					this.m_oLvRowColors.AddRow();
					this.m_oLvRowColors.AddColumns(listView1.Items.Count-1,listView1.Columns.Count);

					listView1.Items[listView1.Items.Count-1].UseItemStyleForSubItems=false;

					this.m_oLvRowColors.ListViewSubItem(listView1.Items[listView1.Items.Count-1].Index,0,listView1.Items[listView1.Items.Count-1].SubItems[listView1.Items[listView1.Items.Count-1].SubItems.Count-1],false);

					
					this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(this.getViewMenuText(strCategory.Trim().ToUpper()));
					this.m_oLvRowColors.ListViewSubItem(listView1.Items[listView1.Items.Count-1].Index,1,listView1.Items[listView1.Items.Count-1].SubItems[listView1.Items[listView1.Items.Count-1].SubItems.Count-1],false);
					this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(strDocument);
					this.m_oLvRowColors.ListViewSubItem(listView1.Items[listView1.Items.Count-1].Index,2,listView1.Items[listView1.Items.Count-1].SubItems[listView1.Items[listView1.Items.Count-1].SubItems.Count-1],false);
					this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(strDesc);		
					this.m_oLvRowColors.ListViewSubItem(listView1.Items[listView1.Items.Count-1].Index,3,listView1.Items[listView1.Items.Count-1].SubItems[listView1.Items[listView1.Items.Count-1].SubItems.Count-1],false);


				}


				this.listView1.Columns[1].Width = -1;
				this.listView1.Columns[2].Width = -1;
				this.listView1.Columns[3].Width = -1;
				this.UpdateListViewCount();
			}


		}

		public void savevalues(string streditmode,string strDocument, string strDesc,string strCategory, string strSharedPrivate)
		{
			
			string strSQL="";
			string strSourceFile="";
			string strConn="";

            ado_data_access p_ado = new ado_data_access();  
			switch (streditmode)
			{
				case "n": //new
					strSQL = "INSERT INTO links_depository (category,subcategory,link,description,list_yn) VALUES " + "(" +  
						     Convert.ToString(this.getCategoryAssignment(strCategory)) + "," + 
						     Convert.ToString(this.getSubCategoryAssignment(strCategory)) + "," + 
						"'" + strDocument +  "'," + 
						"'" + strDesc + "'," + 
						"'Y');";
                    break;
				case "e": //edit
					strSQL = "UPDATE links_depository SET category = " + Convert.ToString(this.getCategoryAssignment(strCategory)) + ", " +
						"subcategory = " +  Convert.ToString(this.getSubCategoryAssignment(strCategory)) + ", " +
						"link = '" + strDocument + "', " + 
						"description = '" + strDesc + "' WHERE category=" +
						Convert.ToString(this.getCategoryAssignment(this.listView1.SelectedItems[0].SubItems[1].Text)) + 
						" AND subcategory=" + Convert.ToString(this.getSubCategoryAssignment(this.listView1.SelectedItems[0].SubItems[1].Text)) + 
						" AND link='" + this.listView1.SelectedItems[0].SubItems[2].Text + "';";
					break;
				case "d":   //delete
					strSQL = "DELETE FROM links_depository WHERE category=" +
						Convert.ToString(this.getCategoryAssignment(this.listView1.SelectedItems[0].SubItems[1].Text)) + 
						" AND subcategory=" + Convert.ToString(this.getSubCategoryAssignment(this.listView1.SelectedItems[0].SubItems[1].Text)) + 
						" AND link='" + this.listView1.SelectedItems[0].SubItems[2].Text + "';";
					break;

			}
			
			if (strSharedPrivate == "S" || strSharedPrivate == "B")
			{
			    strSourceFile = ((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text.Trim() + "\\shared_project_links_and_notes.mdb";
				strConn=p_ado.getMDBConnString(strSourceFile,"admin","");
				//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strSourceFile + ";User Id=admin;Password=;";
				p_ado.SqlNonQuery(strConn,strSQL);
			}

			if (strSharedPrivate == "P" || strSharedPrivate == "B")
			{
				strSourceFile = ((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text.Trim() + "\\personal_project_links_and_notes.mdb";
				strConn=p_ado.getMDBConnString(strSourceFile,"admin","");
				//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strSourceFile + ";User Id=admin;Password=;";
				p_ado.SqlNonQuery(strConn,strSQL);
			}
            this.m_intError=p_ado.m_intError;
			p_ado = null;

			


		}
		public void loadvalues(string strSharedDir, string strPrivateDir, bool lAlwaysLoad)
		{

			//string strSQL="";
			//int intArrayCount;
			//int x=0;
			//string strConn="";
			//string strCommand="";
			//string str="";
			//string strMDB="";

			this.m_oLvRowColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
			this.m_oLvRowColors.ReferenceForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvRowColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
			this.m_oLvRowColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvRowColors.ReferenceSelectedRowBackgroundColor=frmMain.g_oGridViewSelectedRowBackgroundColor;
			this.m_oLvRowColors.CustomFullRowSelect=true;
			this.m_oLvRowColors.ReferenceListView = this.listView1;
			if (frmMain.g_oGridViewFont != null) this.listView1.Font = frmMain.g_oGridViewFont;

            

			//see if this is the first time loading the records
			if (this.sbDisplayedRecordCount.Text.Trim().Length == 0 || lAlwaysLoad==true)
			{  
				this.listView1.Clear();
				this.m_oLvRowColors.InitializeRowCollection();
				this.listView1.Columns.Add(" ",this.imageList1.Images[0].Width,HorizontalAlignment.Left);
				this.listView1.Columns.Add(" ", 30 , HorizontalAlignment.Center);
				this.listView1.Columns.Add("Document", 250, HorizontalAlignment.Left);
				this.listView1.Columns.Add("Description", 500, HorizontalAlignment.Left);
				string strSQLWhere = this.getViewMenuSQL();
				if (this.btnBoth.Pushed==true || this.btnGlobal.Pushed == true)
				{
					if (strSharedDir.Trim().Length > 0)
					{
						
						this.loadsharedlinks(strSharedDir,strSQLWhere);
					}
					else 
					{

					}
				}
                
				if (this.btnBoth.Pushed==true || this.btnPrivate.Pushed==true)
				{
					if (strPrivateDir.Trim().Length > 0)
					{
						
						this.loadprivatelinks(strPrivateDir,strSQLWhere);
					}
					else 
					{

					}

				}
                
				if (this.btnPrivate.Enabled==false && this.btnGlobal.Enabled==false)
				{
				}
				else 
				{
					if (this.btnPrivate.Enabled==false || this.btnGlobal.Enabled==false)
					{
					}   
					else 
					{

					}

					if (this.listView1.Items.Count > 0) 
					{
						this.listView1.Columns[0].Width = -1;
						this.listView1.Columns[1].Width = -1;
					}

				}
			}
            this.UpdateListViewCount();
		}
		private void UpdateListViewCount()
		{
			try
			{
				if (this.listView1.Items.Count > 0)
				{
                    if (this.listView1.SelectedItems.Count == 0)
                    {
                        this.listView1.Items[0].Selected = true;
                        System.Threading.Thread.Sleep(1000);
                    }
					this.sbQueryRecordCount.Text = Convert.ToString(this.listView1.SelectedItems[0].Index + 1) + "/" + this.listView1.Items.Count;
					this.sbDisplayedRecordCount.Text = this.listView1.Items.Count.ToString().Trim();            
				}
				else 
				{
					this.sbQueryRecordCount.Text = "0/0";
					this.sbDisplayedRecordCount.Text = "0";
				}
			}
			catch (Exception caught)
			{
				MessageBox.Show(caught.ToString());
			}

		}

		private void listView1_SelectedIndexChanged(object sender, System.EventArgs e)
		{

            if (this.listView1.SelectedItems.Count == 0) return; 
			this.sbQueryRecordCount.Text = Convert.ToString(this.listView1.SelectedItems[0].Index + 1) + "/" + this.listView1.Items.Count;
			this.m_oLvRowColors.DelegateListViewItem(listView1.SelectedItems[0]);
		}
		private bool RemoveFromCurrentList(string strCategory)
		{
			const byte ALL = 0;
			//const byte CORE = 1;
			//const byte FVS = 2;
			//const byte FRCS = 3;
			//const byte PROCESSOR = 4;
			//const byte GIS = 5;
			//const byte OTHER = 6;
            int x=0;
			//check if viewing all categories
			if (this.contextMenu1.MenuItems[ALL].Checked == true) return false;


            
			//string strMenuItem = this.getViewMenuText(strCategory);

			for (x=0; x<= this.contextMenu1.MenuItems.Count-1;x++) 
			{
				if (this.contextMenu1.MenuItems[x].Checked == true &&
					this.contextMenu1.MenuItems[x].Text.Trim().ToUpper() == 
					strCategory.Trim().ToUpper()) 
				{
					return false;
				}
			}
			return true;
		}

		private void listView1_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			
			try
			{
				if (e.Button == MouseButtons.Left)
				{
					int intRowHt = listView1.Items[0].Bounds.Height;
					double dblRow = (double)(e.Y / intRowHt);
					this.listView1.Items[listView1.TopItem.Index + (int)dblRow-1].Selected=true;
					
				}
			}
			catch 
			{
			}
		}
	}
}
