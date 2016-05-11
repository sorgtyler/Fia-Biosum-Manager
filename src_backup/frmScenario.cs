using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class frmScenario : System.Windows.Forms.Form
	{
		public System.Windows.Forms.TextBox txtDropDown;
		private System.Data.DataView dataView1;
		public string strScenarioDescription;
		private System.Windows.Forms.ImageList imgSize;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.ImageList imageList2;
		public bool m_bScenarioOpen = false;
		public bool m_ldatasourcefirsttime;
		public bool m_lrulesfirsttime=true;
		private System.Windows.Forms.Button btnClose;
		public FIA_Biosum_Manager.uc_scenario uc_scenario1;
		public FIA_Biosum_Manager.uc_datasource uc_datasource1;
		public FIA_Biosum_Manager.uc_scenario_ffe uc_scenario_ffe1;
		public FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_effective uc_scenario_fvs_prepost_variables_effective1;
		public FIA_Biosum_Manager.uc_scenario_fvs_prepost_optimization uc_scenario_fvs_prepost_optimization1;
		public FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_tiebreaker uc_scenario_fvs_prepost_variables_tiebreaker1;
		//public FIA_Biosum_Manager.uc_scenario_treatment_intensity uc_scenario_treatment_intensity1;
		public FIA_Biosum_Manager.uc_scenario_costs uc_scenario_costs1;
		public FIA_Biosum_Manager.uc_select_list_item uc_select_list_item1;
		public FIA_Biosum_Manager.uc_scenario_filter uc_scenario_filter1;
		public FIA_Biosum_Manager.uc_scenario_filter uc_scenario_cond_filter1;
		public FIA_Biosum_Manager.uc_scenario_psite uc_scenario_psite1;
		public FIA_Biosum_Manager.uc_scenario_open uc_scenario_open1;
		public FIA_Biosum_Manager.uc_scenario_run uc_scenario_run1;
		public FIA_Biosum_Manager.uc_scenario_notes uc_scenario_notes1;
		public FIA_Biosum_Manager.uc_scenario_owner_groups uc_scenario_owner_groups1;
		public FIA_Biosum_Manager.frmGridView frmGridView1;
		private FIA_Biosum_Manager.frmRunCoreScenario frmRunCoreScenario1;
		
		public System.Windows.Forms.VScrollBar m_vScrollBar;
		public System.Windows.Forms.HScrollBar m_hScrollBar;
		protected int m_intIntensityTop=0;
		protected int m_intLblTitleTop = 0;
		private int m_intHScrollValue=0;
		private double m_dblHScrollOldPerc=0;
		private double m_dblHScrollNewPerc=0;
		private int m_intHScrollMaxSize=0;
		public System.Windows.Forms.TabControl tabControlScenario;
		private System.Windows.Forms.TabPage tbDesc;
		private System.Windows.Forms.TabPage tbNotes;
		private System.Windows.Forms.TabPage tbDataSources;
		private System.Windows.Forms.TabPage tbRules;
		public System.Windows.Forms.ToolBar tlbScenario;
		public System.Windows.Forms.TabControl tabControlRules;
		private System.Windows.Forms.TabPage tbOwners;
		private System.Windows.Forms.TabPage tbPSites;
		private System.Windows.Forms.TabPage tbFilterPlots;
		private System.Windows.Forms.ToolBarButton btnScenarioNew;
		private System.Windows.Forms.ToolBarButton btnScenarioOpen;
		private System.Windows.Forms.ToolBarButton btnScenarioSave;
		private System.Windows.Forms.ToolBarButton toolBarButton1;
		private System.Windows.Forms.ToolBarButton btnScenarioDelete;
		private System.Windows.Forms.TabPage tbRun;
		public bool m_bSave=false;
		private System.Windows.Forms.TabPage tbCost;
		private System.Windows.Forms.TabPage tbFVSVariables;
		public System.Windows.Forms.TabPage tbOptimization;
		public bool m_bOptimizeTabPageEnabled=true;

		private bool m_oTabPageLast;


		public string m_strAllowLeaveTabPageMsg="";
		public bool m_bEnableSelectedTabPage=true;
		public int  m_intEditTabPageIndex=0;
		public string m_strCurrentEditTabControlName="";
		public string m_strCurrentEditTabPageText="";
		public System.Windows.Forms.TabControl tabControlFVSVariables;
		private System.Windows.Forms.TabPage tbTieBreaker;
		public System.Windows.Forms.TabPage tbEffective;
		private System.Windows.Forms.TabPage tbFilterCond;
		
		
		public bool m_bPopup = false;

		
		public frmScenario(FIA_Biosum_Manager.frmMain p_frmMain)
		{
			//
			// Required for Windows Form Designer support
			//

			FIA_Biosum_Manager.frmMain.g_oFrmMain=p_frmMain;

			InitializeComponent();
			
			if (this.Width > p_frmMain.Width - p_frmMain.grpboxLeft.Width ) 
			{
				this.Width = p_frmMain.Width - p_frmMain.grpboxLeft.Width  - 10;
			}
			try
			{
			
				this.uc_select_list_item1 = new uc_select_list_item();
				this.frmGridView1 = new frmGridView();
				//
				//scenario description and properties
				//
				this.uc_scenario1 = new uc_scenario();
				
				this.tbDesc.Controls.Add(uc_scenario1);
				uc_scenario1.btnCancel.Hide();
				uc_scenario1.btnClose.Hide();
				uc_scenario1.btnOpen.Hide();
				uc_scenario1.txtScenarioId.Enabled=false;
				uc_scenario1.txtScenarioPath.Enabled=false;
				uc_scenario1.txtDescription.Enabled=true;
				uc_scenario1.lblTitle.Text = "Scenario";
				uc_scenario1.Dock = System.Windows.Forms.DockStyle.Fill;
				uc_scenario1.ReferenceCoreScenarioForm=this;
				uc_scenario1.ScenarioType="core";
				//
				//scenario datasource
				//
				this.uc_datasource1 = new uc_datasource();
				this.tbDataSources.Controls.Add(uc_datasource1);
				uc_datasource1.Dock = System.Windows.Forms.DockStyle.Fill;
				uc_datasource1.ReferenceCoreScenarioForm=this;
				uc_datasource1.ScenarioType="core";
				//
				//scenario notes
				//
				this.uc_scenario_notes1 = new uc_scenario_notes();
				this.tbNotes.Controls.Add(uc_scenario_notes1);
				uc_scenario_notes1.Dock = System.Windows.Forms.DockStyle.Fill;
				uc_scenario_notes1.ReferenceCoreScenarioForm=this;
				uc_scenario_notes1.ScenarioType="core";
				//
				//rule definitions owner groups
				//
				this.uc_scenario_owner_groups1 = new uc_scenario_owner_groups();
				this.tbOwners.Controls.Add(uc_scenario_owner_groups1);
				this.uc_scenario_owner_groups1.Dock = System.Windows.Forms.DockStyle.Fill;
				uc_scenario_owner_groups1.ReferenceCoreScenarioForm=this;
				

				//
				//rule definitions costs
				//
				this.uc_scenario_costs1 = new uc_scenario_costs();
				this.tbCost.Controls.Add(uc_scenario_costs1);
				this.uc_scenario_costs1.Dock = System.Windows.Forms.DockStyle.Fill;
				this.uc_scenario_costs1.ReferenceCoreScenarioForm=this;
				//
				//rule definitions processing sites
				//
				this.uc_scenario_psite1 = new uc_scenario_psite();
				this.tbPSites.Controls.Add(uc_scenario_psite1);
				this.uc_scenario_psite1.Dock = System.Windows.Forms.DockStyle.Fill;
				this.uc_scenario_psite1.ReferenceCoreScenarioForm=this;
				//
				//rule custom plot filter
				//
				this.uc_scenario_filter1 = new uc_scenario_filter();
				this.tbFilterPlots.Controls.Add(uc_scenario_filter1);
				this.uc_scenario_filter1.Dock = System.Windows.Forms.DockStyle.Fill;
				this.uc_scenario_filter1.ReferenceCoreScenarioForm=this;
				this.uc_scenario_filter1.FilterType="PLOT";
				//
				//rule custom condition filter
				//
				this.uc_scenario_cond_filter1 = new uc_scenario_filter();
				this.tbFilterCond.Controls.Add(uc_scenario_cond_filter1);
				this.uc_scenario_cond_filter1.Dock = System.Windows.Forms.DockStyle.Fill;
				this.uc_scenario_cond_filter1.ReferenceCoreScenarioForm=this;
				this.uc_scenario_cond_filter1.lblTitle.Text = "Condition Filter";
				this.uc_scenario_cond_filter1.FilterType="COND";

				
				//
				//rule definitions fvs pre-post variables
				//
				this.uc_scenario_fvs_prepost_variables_effective1 = new uc_scenario_fvs_prepost_variables_effective();
				this.tbEffective.Controls.Add(this.uc_scenario_fvs_prepost_variables_effective1);
				this.uc_scenario_fvs_prepost_variables_effective1.Dock = System.Windows.Forms.DockStyle.Fill;
				this.uc_scenario_fvs_prepost_variables_effective1.ReferenceCoreScenarioForm=this;
				//
				//rule definitions fvs optimization variables
				//
				this.uc_scenario_fvs_prepost_optimization1 = new uc_scenario_fvs_prepost_optimization();
				this.tbOptimization.Controls.Add(this.uc_scenario_fvs_prepost_optimization1);
				this.uc_scenario_fvs_prepost_optimization1.Dock = System.Windows.Forms.DockStyle.Fill;
				this.uc_scenario_fvs_prepost_optimization1.ReferenceCoreScenarioForm = this;
				this.uc_scenario_fvs_prepost_variables_effective1.ReferenceOptimizationUserControl=this.uc_scenario_fvs_prepost_optimization1;

				//
				//rule definitions fvs tie breaker
				//
				this.uc_scenario_fvs_prepost_variables_tiebreaker1 = new FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_tiebreaker();
				this.tbTieBreaker.Controls.Add(this.uc_scenario_fvs_prepost_variables_tiebreaker1);
				this.uc_scenario_fvs_prepost_variables_tiebreaker1.Dock = System.Windows.Forms.DockStyle.Fill;
				this.uc_scenario_fvs_prepost_variables_tiebreaker1.ReferenceCoreScenarioForm = this;
				this.uc_scenario_fvs_prepost_variables_effective1.ReferenceTieBreakerUserControl=this.uc_scenario_fvs_prepost_variables_tiebreaker1;
				this.uc_scenario_fvs_prepost_variables_tiebreaker1.uc_scenario_treatment_intensity1.ReferenceCoreScenarioForm=this;
				this.uc_scenario_fvs_prepost_variables_tiebreaker1.ReferenceOptimizationUserControl=this.uc_scenario_fvs_prepost_optimization1;
				this.uc_scenario_fvs_prepost_optimization1.ReferenceTieBreaker=this.uc_scenario_fvs_prepost_variables_tiebreaker1;


				//rule execute run
				this.uc_scenario_run1 = new uc_scenario_run();
				this.tbRun.Controls.Add(uc_scenario_run1);
				this.uc_scenario_run1.Dock = System.Windows.Forms.DockStyle.Fill;
				this.uc_scenario_run1.ReferenceCoreScenarioForm=this;
				this.uc_scenario_run1.ReferenceFVSPrePostVariables=this.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar;
				this.uc_scenario_run1.ReferenceFVSPrePostOptimization=this.uc_scenario_fvs_prepost_optimization1.m_oSavVariableCollection;
				this.uc_scenario_run1.ReferenceFVSPrePostTieBreaker = this.uc_scenario_fvs_prepost_variables_tiebreaker1.m_oSavTieBreakerCollection;
				this.btnClose.Enabled=true;
				this.resize_frmScenario();



			}
			catch (Exception p_msg)
			{
				MessageBox.Show(p_msg.Message);
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

		public frmScenario()
		{
			this.InitializeComponent();
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmScenario));
			this.tlbScenario = new System.Windows.Forms.ToolBar();
			this.btnScenarioNew = new System.Windows.Forms.ToolBarButton();
			this.btnScenarioOpen = new System.Windows.Forms.ToolBarButton();
			this.toolBarButton1 = new System.Windows.Forms.ToolBarButton();
			this.btnScenarioSave = new System.Windows.Forms.ToolBarButton();
			this.btnScenarioDelete = new System.Windows.Forms.ToolBarButton();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.txtDropDown = new System.Windows.Forms.TextBox();
			this.dataView1 = new System.Data.DataView();
			this.imgSize = new System.Windows.Forms.ImageList(this.components);
			this.imageList2 = new System.Windows.Forms.ImageList(this.components);
			this.btnClose = new System.Windows.Forms.Button();
			this.tabControlScenario = new System.Windows.Forms.TabControl();
			this.tbDesc = new System.Windows.Forms.TabPage();
			this.tbNotes = new System.Windows.Forms.TabPage();
			this.tbDataSources = new System.Windows.Forms.TabPage();
			this.tbRules = new System.Windows.Forms.TabPage();
			this.tabControlRules = new System.Windows.Forms.TabControl();
			this.tbOwners = new System.Windows.Forms.TabPage();
			this.tbCost = new System.Windows.Forms.TabPage();
			this.tbPSites = new System.Windows.Forms.TabPage();
			this.tbFilterPlots = new System.Windows.Forms.TabPage();
			this.tbFilterCond = new System.Windows.Forms.TabPage();
			this.tbFVSVariables = new System.Windows.Forms.TabPage();
			this.tabControlFVSVariables = new System.Windows.Forms.TabControl();
			this.tbEffective = new System.Windows.Forms.TabPage();
			this.tbOptimization = new System.Windows.Forms.TabPage();
			this.tbTieBreaker = new System.Windows.Forms.TabPage();
			this.tbRun = new System.Windows.Forms.TabPage();
			((System.ComponentModel.ISupportInitialize)(this.dataView1)).BeginInit();
			this.tabControlScenario.SuspendLayout();
			this.tbRules.SuspendLayout();
			this.tabControlRules.SuspendLayout();
			this.tbFVSVariables.SuspendLayout();
			this.tabControlFVSVariables.SuspendLayout();
			this.SuspendLayout();
			// 
			// tlbScenario
			// 
			this.tlbScenario.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
																						   this.btnScenarioNew,
																						   this.btnScenarioOpen,
																						   this.toolBarButton1,
																						   this.btnScenarioSave,
																						   this.btnScenarioDelete});
			this.tlbScenario.DropDownArrows = true;
			this.tlbScenario.ImageList = this.imageList1;
			this.tlbScenario.Location = new System.Drawing.Point(0, 0);
			this.tlbScenario.Name = "tlbScenario";
			this.tlbScenario.ShowToolTips = true;
			this.tlbScenario.Size = new System.Drawing.Size(663, 28);
			this.tlbScenario.TabIndex = 42;
			this.tlbScenario.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.tlbScenario_ButtonClick);
			// 
			// btnScenarioNew
			// 
			this.btnScenarioNew.ImageIndex = 1;
			this.btnScenarioNew.ToolTipText = "New";
			// 
			// btnScenarioOpen
			// 
			this.btnScenarioOpen.ImageIndex = 0;
			this.btnScenarioOpen.ToolTipText = "Open";
			// 
			// toolBarButton1
			// 
			this.toolBarButton1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
			// 
			// btnScenarioSave
			// 
			this.btnScenarioSave.ImageIndex = 2;
			this.btnScenarioSave.ToolTipText = "Save";
			// 
			// btnScenarioDelete
			// 
			this.btnScenarioDelete.ImageIndex = 3;
			this.btnScenarioDelete.ToolTipText = "Delete";
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// txtDropDown
			// 
			this.txtDropDown.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
			this.txtDropDown.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDropDown.ForeColor = System.Drawing.SystemColors.ControlText;
			this.txtDropDown.HideSelection = false;
			this.txtDropDown.Location = new System.Drawing.Point(224, 416);
			this.txtDropDown.Multiline = true;
			this.txtDropDown.Name = "txtDropDown";
			this.txtDropDown.ReadOnly = true;
			this.txtDropDown.Size = new System.Drawing.Size(24, 24);
			this.txtDropDown.TabIndex = 12;
			this.txtDropDown.Text = "";
			this.txtDropDown.Visible = false;
			// 
			// imgSize
			// 
			this.imgSize.ImageSize = new System.Drawing.Size(16, 16);
			this.imgSize.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imgSize.ImageStream")));
			this.imgSize.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// imageList2
			// 
			this.imageList2.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit;
			this.imageList2.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList2.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList2.ImageStream")));
			this.imageList2.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btnClose
			// 
			this.btnClose.Enabled = false;
			this.btnClose.Location = new System.Drawing.Point(560, 416);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(96, 32);
			this.btnClose.TabIndex = 40;
			this.btnClose.Text = "Close";
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// tabControlScenario
			// 
			this.tabControlScenario.Controls.Add(this.tbDesc);
			this.tabControlScenario.Controls.Add(this.tbNotes);
			this.tabControlScenario.Controls.Add(this.tbDataSources);
			this.tabControlScenario.Controls.Add(this.tbRules);
			this.tabControlScenario.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
			this.tabControlScenario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.tabControlScenario.ItemSize = new System.Drawing.Size(100, 18);
			this.tabControlScenario.Location = new System.Drawing.Point(0, 32);
			this.tabControlScenario.Name = "tabControlScenario";
			this.tabControlScenario.SelectedIndex = 0;
			this.tabControlScenario.Size = new System.Drawing.Size(648, 360);
			this.tabControlScenario.TabIndex = 41;
			this.tabControlScenario.TabIndexChanged += new System.EventHandler(this.tabControlScenario_TabIndexChanged);
			this.tabControlScenario.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.tabControlScenario_DrawItem);
			this.tabControlScenario.Enter += new System.EventHandler(this.tabControlScenario_Enter);
			this.tabControlScenario.Leave += new System.EventHandler(this.tabControlScenario_Leave);
			this.tabControlScenario.ChangeUICues += new System.Windows.Forms.UICuesEventHandler(this.tabControlScenario_ChangeUICues);
			this.tabControlScenario.SelectedIndexChanged += new System.EventHandler(this.tabControlScenario_SelectedIndexChanged);
			// 
			// tbDesc
			// 
			this.tbDesc.AutoScroll = true;
			this.tbDesc.Location = new System.Drawing.Point(4, 22);
			this.tbDesc.Name = "tbDesc";
			this.tbDesc.Size = new System.Drawing.Size(640, 334);
			this.tbDesc.TabIndex = 0;
			this.tbDesc.Text = "Description";
			// 
			// tbNotes
			// 
			this.tbNotes.AutoScroll = true;
			this.tbNotes.Location = new System.Drawing.Point(4, 22);
			this.tbNotes.Name = "tbNotes";
			this.tbNotes.Size = new System.Drawing.Size(640, 334);
			this.tbNotes.TabIndex = 1;
			this.tbNotes.Text = "Notes";
			// 
			// tbDataSources
			// 
			this.tbDataSources.AutoScroll = true;
			this.tbDataSources.Location = new System.Drawing.Point(4, 22);
			this.tbDataSources.Name = "tbDataSources";
			this.tbDataSources.Size = new System.Drawing.Size(640, 334);
			this.tbDataSources.TabIndex = 2;
			this.tbDataSources.Text = "Data Sources";
			// 
			// tbRules
			// 
			this.tbRules.Controls.Add(this.tabControlRules);
			this.tbRules.ForeColor = System.Drawing.Color.Red;
			this.tbRules.Location = new System.Drawing.Point(4, 22);
			this.tbRules.Name = "tbRules";
			this.tbRules.Size = new System.Drawing.Size(640, 334);
			this.tbRules.TabIndex = 3;
			this.tbRules.Text = "Rule Definitions";
			this.tbRules.Click += new System.EventHandler(this.tbRules_Click);
			// 
			// tabControlRules
			// 
			this.tabControlRules.Controls.Add(this.tbOwners);
			this.tabControlRules.Controls.Add(this.tbCost);
			this.tabControlRules.Controls.Add(this.tbPSites);
			this.tabControlRules.Controls.Add(this.tbFilterPlots);
			this.tabControlRules.Controls.Add(this.tbFilterCond);
			this.tabControlRules.Controls.Add(this.tbFVSVariables);
			this.tabControlRules.Controls.Add(this.tbRun);
			this.tabControlRules.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tabControlRules.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
			this.tabControlRules.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.tabControlRules.Location = new System.Drawing.Point(0, 0);
			this.tabControlRules.Name = "tabControlRules";
			this.tabControlRules.SelectedIndex = 0;
			this.tabControlRules.Size = new System.Drawing.Size(640, 334);
			this.tabControlRules.TabIndex = 0;
			this.tabControlRules.TabIndexChanged += new System.EventHandler(this.tabControlRules_TabIndexChanged);
			this.tabControlRules.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.tabControlRules_DrawItem);
			this.tabControlRules.SelectedIndexChanged += new System.EventHandler(this.tabControlRules_SelectedIndexChanged);
			// 
			// tbOwners
			// 
			this.tbOwners.Location = new System.Drawing.Point(4, 22);
			this.tbOwners.Name = "tbOwners";
			this.tbOwners.Size = new System.Drawing.Size(632, 308);
			this.tbOwners.TabIndex = 1;
			this.tbOwners.Text = "Land Ownership Groups";
			this.tbOwners.Click += new System.EventHandler(this.tbOwners_Click);
			// 
			// tbCost
			// 
			this.tbCost.Location = new System.Drawing.Point(4, 25);
			this.tbCost.Name = "tbCost";
			this.tbCost.Size = new System.Drawing.Size(632, 305);
			this.tbCost.TabIndex = 7;
			this.tbCost.Text = "Cost And Revenue";
			// 
			// tbPSites
			// 
			this.tbPSites.Location = new System.Drawing.Point(4, 25);
			this.tbPSites.Name = "tbPSites";
			this.tbPSites.Size = new System.Drawing.Size(632, 305);
			this.tbPSites.TabIndex = 4;
			this.tbPSites.Text = "Wood Processing Sites";
			this.tbPSites.Click += new System.EventHandler(this.tbPSites_Click);
			this.tbPSites.Enter += new System.EventHandler(this.tbPSites_Enter);
			// 
			// tbFilterPlots
			// 
			this.tbFilterPlots.Location = new System.Drawing.Point(4, 25);
			this.tbFilterPlots.Name = "tbFilterPlots";
			this.tbFilterPlots.Size = new System.Drawing.Size(632, 305);
			this.tbFilterPlots.TabIndex = 5;
			this.tbFilterPlots.Text = "Filter Plot Records";
			// 
			// tbFilterCond
			// 
			this.tbFilterCond.Location = new System.Drawing.Point(4, 25);
			this.tbFilterCond.Name = "tbFilterCond";
			this.tbFilterCond.Size = new System.Drawing.Size(632, 305);
			this.tbFilterCond.TabIndex = 9;
			this.tbFilterCond.Text = "Filter Condition Records";
			// 
			// tbFVSVariables
			// 
			this.tbFVSVariables.Controls.Add(this.tabControlFVSVariables);
			this.tbFVSVariables.Location = new System.Drawing.Point(4, 25);
			this.tbFVSVariables.Name = "tbFVSVariables";
			this.tbFVSVariables.Size = new System.Drawing.Size(632, 305);
			this.tbFVSVariables.TabIndex = 8;
			this.tbFVSVariables.Text = "FVS Variables";
			// 
			// tabControlFVSVariables
			// 
			this.tabControlFVSVariables.Controls.Add(this.tbEffective);
			this.tabControlFVSVariables.Controls.Add(this.tbOptimization);
			this.tabControlFVSVariables.Controls.Add(this.tbTieBreaker);
			this.tabControlFVSVariables.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tabControlFVSVariables.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
			this.tabControlFVSVariables.Location = new System.Drawing.Point(0, 0);
			this.tabControlFVSVariables.Name = "tabControlFVSVariables";
			this.tabControlFVSVariables.SelectedIndex = 0;
			this.tabControlFVSVariables.Size = new System.Drawing.Size(632, 305);
			this.tabControlFVSVariables.TabIndex = 0;
			this.tabControlFVSVariables.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.tbFVSVariablesSelect_DrawItem);
			this.tabControlFVSVariables.SelectedIndexChanged += new System.EventHandler(this.tbFVSVariablesSelect_SelectedIndexChanged);
			// 
			// tbEffective
			// 
			this.tbEffective.Location = new System.Drawing.Point(4, 25);
			this.tbEffective.Name = "tbEffective";
			this.tbEffective.Size = new System.Drawing.Size(624, 279);
			this.tbEffective.TabIndex = 0;
			this.tbEffective.Text = "Effective";
			// 
			// tbOptimization
			// 
			this.tbOptimization.Location = new System.Drawing.Point(4, 22);
			this.tbOptimization.Name = "tbOptimization";
			this.tbOptimization.Size = new System.Drawing.Size(624, 282);
			this.tbOptimization.TabIndex = 1;
			this.tbOptimization.Text = "Optimization";
			// 
			// tbTieBreaker
			// 
			this.tbTieBreaker.Location = new System.Drawing.Point(4, 22);
			this.tbTieBreaker.Name = "tbTieBreaker";
			this.tbTieBreaker.Size = new System.Drawing.Size(624, 282);
			this.tbTieBreaker.TabIndex = 2;
			this.tbTieBreaker.Text = "Tie Breaker";
			// 
			// tbRun
			// 
			this.tbRun.Location = new System.Drawing.Point(4, 25);
			this.tbRun.Name = "tbRun";
			this.tbRun.Size = new System.Drawing.Size(632, 305);
			this.tbRun.TabIndex = 6;
			this.tbRun.Text = "Run";
			// 
			// frmScenario
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.BackColor = System.Drawing.SystemColors.Control;
			this.ClientSize = new System.Drawing.Size(663, 454);
			this.Controls.Add(this.tlbScenario);
			this.Controls.Add(this.tabControlScenario);
			this.Controls.Add(this.btnClose);
			this.Controls.Add(this.txtDropDown);
			this.Name = "frmScenario";
			this.Text = "Case Study Scenario";
			this.Resize += new System.EventHandler(this.frmScenario_Resize);
			this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.frmScenario_MouseDown);
			this.Click += new System.EventHandler(this.frmScenario_Click);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmScenario_Closing);
			this.Load += new System.EventHandler(this.frmscenarioScenario_Load);
			((System.ComponentModel.ISupportInitialize)(this.dataView1)).EndInit();
			this.tabControlScenario.ResumeLayout(false);
			this.tbRules.ResumeLayout(false);
			this.tabControlRules.ResumeLayout(false);
			this.tbFVSVariables.ResumeLayout(false);
			this.tabControlFVSVariables.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void frmscenarioScenario_Load(object sender, System.EventArgs e)
		{
			this.resize_frmScenario();
		}

		private void btnCurrentscenario_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnCurrentscenario_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{

		}

		private void btnCurrentscenario_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			this.txtDropDown.Visible=false;
		}

		private void btnSize_Click(object sender, System.EventArgs e)
		{
		}

		private void frmScenario_Click(object sender, System.EventArgs e)
		{
			this.Focus();
		}

		private void frmScenario_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.resize_frmScenario();
			}
			catch
			{
			}
		}
		public void resize_frmScenario()
		{
			try
			{
				if (this.uc_scenario_open1 !=null)
				{
					this.uc_scenario_open1.Width = this.ClientSize.Width - 2;
					this.uc_scenario_open1.Height = this.ClientSize.Height - this.uc_scenario_open1.Top - 2;
				}

				if (this.uc_scenario_ffe1 !=null)
				{
					if (this.uc_scenario_ffe1.Visible==true)
					{
					}
					else 
					{
					}
				}
				this.btnClose.Top = this.ClientSize.Height - this.btnClose.Height - 2;
				this.btnClose.Left = this.ClientSize.Width - this.btnClose.Width - 2;
				this.tabControlScenario.Top = this.tlbScenario.Top + this.tlbScenario.Height + 2;
				this.tabControlScenario.Height = this.btnClose.Top - this.tabControlScenario.Top - 2;
				this.tabControlScenario.Width = this.ClientSize.Width;

				
			}
			catch
			{
			}
			finally 
			{
			}

		}

		private void btnDataSources_Click(object sender, System.EventArgs e)
		{
		}

		private void mnuFile_Click(object sender, System.EventArgs e)
		{
		
		}

		

		private void btnDataSources_Click_1(object sender, System.EventArgs e)
		{
		
		}

		private void grpboxLeft_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btnFile2_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			
		
		}

		

		
		

		


		private void btnNew_Click(object sender, System.EventArgs e)
		{
			int intHt=0;
			if (this.m_bScenarioOpen == false) 
			{
				this.uc_datasource1.Visible=false;
				this.uc_scenario_notes1.Visible=false;
				this.uc_scenario1.Visible=false;

				this.uc_scenario_owner_groups1.Visible=false;
				this.uc_scenario_costs1.Visible=false;
				this.uc_scenario_psite1.Visible=false;
				this.uc_scenario_filter1.Visible=false;
				this.m_vScrollBar.Visible=false;
				this.m_hScrollBar.Visible=false;
				this.Height = intHt;
				this.Height = intHt + ((intHt - this.ClientSize.Height) * 2);
				this.SetMenu("new");
				this.uc_scenario1.lblTitle.Text = "New Scenario";
				this.uc_scenario1.txtDescription.Enabled=true;
				this.uc_scenario1.resize_uc_scenario();
				this.uc_scenario1.Visible=true;
				
				this.uc_scenario1.NewScenario();
			}
			else 
			{
				frmScenario frmTemp = new frmScenario((frmMain)this.ParentForm);
				frmTemp.Visible=false;
				frmTemp.uc_datasource1.Visible = false;
				frmTemp.uc_scenario_notes1.Visible=false;
				frmTemp.uc_scenario1.Visible=false;
//				frmTemp.uc_scenario_treatment_intensity1.Visible=false;
				frmTemp.uc_scenario_owner_groups1.Visible=false;
				frmTemp.uc_scenario_costs1.Visible=false;
				frmTemp.uc_scenario_psite1.Visible=false;
				frmTemp.uc_scenario_filter1.Visible=false;
				frmTemp.Height = intHt;
				frmTemp.Height = intHt + ((intHt - frmTemp.ClientSize.Height) * 2);
				frmTemp.SetMenu("new");
				frmTemp.BackColor = System.Drawing.SystemColors.Control;
				frmTemp.Text = "Core Analysis: Case Study Scenario";
				frmTemp.MdiParent = this.ParentForm;
				frmTemp.uc_scenario1.lblTitle.Text="New Scenario";
				frmTemp.uc_scenario1.txtDescription.Enabled=true;
				frmTemp.Visible=true;
				frmTemp.uc_scenario1.resize_uc_scenario();
				frmTemp.uc_scenario1.Visible=true;
				frmTemp.uc_scenario1.txtScenarioId.Visible=false;
				frmTemp.uc_scenario1.lblNewScenario.Visible=false;
				frmTemp.uc_scenario1.NewScenario();
				frmTemp.Show();
			}
		}
		public void InitializeNewScenario()
		{
			this.uc_scenario1 = new uc_scenario();

			this.Controls.Add(uc_scenario1);

			uc_scenario1.Dock = System.Windows.Forms.DockStyle.Fill;

			this.tlbScenario.Hide();

			this.tabControlScenario.Hide();

			this.btnClose.Hide();

			this.uc_scenario1.ReferenceCoreScenarioForm=this;

			this.uc_scenario1.ScenarioType="core";


			this.uc_scenario1.NewScenario();

			


		}

		public void InitializeOpenScenario()
		{
			this.uc_scenario_open1 = new uc_scenario_open(); 

			this.Controls.Add(uc_scenario_open1);


			uc_scenario_open1.Left = 0;
			uc_scenario_open1.Width = this.ClientSize.Width;
			uc_scenario_open1.Top = this.tlbScenario.Top + this.tlbScenario.Height;
    		
			this.btnScenarioSave.Enabled=false;
			this.btnScenarioOpen.Enabled=false;

			this.tabControlScenario.Hide();

			this.btnClose.Hide();

			this.uc_scenario_open1.OpenScenario();

			this.Height = 200;
			int intHt = this.Height;
			int intHt2=this.uc_scenario_open1.btnOpen.Height;
			int intTop=this.uc_scenario_open1.btnOpen.Top;
			while (intTop + intHt2 + 20
				>=  intHt)
			{
				intHt += 10;

			}
			this.Height = intHt;

			


		}

		public void SetMenu(string strType)
		{
			switch (strType)
			{
				case "default" :
					break;
				case "new":
					break;
				case "open":
					break;
				case "scenario":
					break;
			}
		}
		private void btnOpen_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
		
		}

		private void btnCurrentScenario_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			string strTemp = this.uc_scenario1.txtDescription.Text.Trim();
			int lines = (int) (strTemp.Length / 30);  //40 characters per line
			if (lines == 0) lines = 1;

			this.txtDropDown.Text = strTemp;
			int textWidth = (int)this.CreateGraphics().MeasureString(strTemp, this.txtDropDown.Font).Width;
			int textHeight = (int)this.CreateGraphics().MeasureString(strTemp, this.txtDropDown.Font).Height;
			this.txtDropDown.Width = (int) ((textWidth / lines) * 1.5);
			this.txtDropDown.Height = (int)Math.Round((textHeight * lines) * 1.5,0) ;
			this.txtDropDown.BringToFront();
			this.txtDropDown.Visible = true;

		}

		private void btnCurrentScenario_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			this.txtDropDown.Visible=false;
		}

		private void btnDelete_Click(object sender, System.EventArgs e)
		{
			this.uc_scenario1.DeleteScenario();
		}

		private void btnOpen_Click(object sender, System.EventArgs e)
		{
			int intAvailWd = this.ParentForm.ClientSize.Width - ((frmMain)this.ParentForm).grpboxLeft.Left - ((frmMain)this.ParentForm).grpboxLeft.Width - 20;
			int intAvailHt = this.ParentForm.ClientSize.Height - ((frmMain)this.ParentForm).tlbMain.Top - ((frmMain)this.ParentForm).tlbMain.Height - 20;

			
			if (this.m_bScenarioOpen == false) 
			{
				this.uc_datasource1.Visible=false;
				this.uc_scenario_notes1.Visible=false;

				this.uc_scenario_owner_groups1.Visible=false;
				this.uc_scenario_costs1.Visible=false;
				this.uc_scenario_psite1.Visible=false;
				this.uc_scenario_filter1.Visible=false;
				this.m_vScrollBar.Visible=false;
				this.m_hScrollBar.Visible=false;


				this.Height = intAvailHt; 
				this.Width = intAvailWd; 
				this.Left = 0;
				this.Top = 0;
				this.uc_scenario1.lblTitle.Text = "Open Scenario";

				this.uc_scenario1.txtScenarioId.Visible=false;
				this.uc_scenario1.lblNewScenario.Visible=false;
				this.uc_scenario1.resize_uc_scenario();
				this.uc_scenario1.Visible=true;
		
			}
			else 
			{
				frmScenario frmTemp = new frmScenario((frmMain)this.ParentForm);
				frmTemp.m_vScrollBar.Visible=false;
				frmTemp.m_hScrollBar.Visible=false;
				frmTemp.BackColor = System.Drawing.SystemColors.Control;
				frmTemp.Text = "Core Analysis: Case Study Scenario";
				frmTemp.MdiParent = this.ParentForm;

				frmTemp.SetMenu("new");
				frmTemp.uc_scenario1.lblTitle.Text = "Open Scenario";
				frmTemp.uc_scenario1.lblTitle.Width =
					(int)this.CreateGraphics().MeasureString(frmTemp.uc_scenario1.lblTitle.Text,frmTemp.uc_scenario1.lblTitle.Font).Width;
				frmTemp.uc_scenario1.txtScenarioId.Visible=false;
				frmTemp.uc_scenario1.lblNewScenario.Visible=false;
				frmTemp.uc_scenario1.resize_uc_scenario();
				frmTemp.uc_scenario1.Visible=true;
				frmTemp.Show();
				frmTemp.Height = intAvailHt; 
				frmTemp.Width = intAvailWd; 
				frmTemp.Left = 0;
				frmTemp.Top = 0;
			}

		}

		private void contextMenuView_Popup(object sender, System.EventArgs e)
		{
			m_bPopup=true;
		}

		private void frmScenario_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (m_bPopup == true) m_bPopup =false;
		}

		public void RulesRepositionControls()
		{
			try
			{
				this.uc_scenario_ffe1.Top = this.uc_scenario_owner_groups1.Top + this.uc_scenario_owner_groups1.Height;
				this.uc_scenario_costs1.Top = this.uc_scenario_ffe1.Top + this.uc_scenario_ffe1.Height;			  
				this.uc_scenario_psite1.Top = this.uc_scenario_costs1.Top + this.uc_scenario_costs1.Height;
				this.uc_scenario_filter1.Top = this.uc_scenario_psite1.Top + this.uc_scenario_psite1.Height;
				this.btnClose.Top = this.uc_scenario_filter1.Top + 
					this.uc_scenario_filter1.Height + 5;
				if (this.btnClose.Top + this.btnClose.Height + 300 > this.m_vScrollBar.Maximum)
					this.m_vScrollBar.Maximum = this.btnClose.Top + this.btnClose.Height + 300;
			}
			catch
			{
			}
		}

		private void btnSave_Click(object sender, System.EventArgs e)
		{
			SaveRuleDefinitions();


		}
		private void LoadRuleDefinitions()
		{
			frmTherm p_frmTherm = new frmTherm();
			p_frmTherm.lblMsg.Text = "";
			p_frmTherm.progressBar1.Minimum = 0;
			p_frmTherm.progressBar1.Maximum = 8;
			p_frmTherm.btnCancel.Visible=false;
			p_frmTherm.lblMsg.Visible=true;
			p_frmTherm.Show();
			p_frmTherm.progressBar1.Value=1;
			p_frmTherm.progressBar1.Value=2;

			p_frmTherm.progressBar1.Value=3;
			p_frmTherm.lblMsg.Text = "Rule Definitions: FVS Variables Data";
			p_frmTherm.lblMsg.Refresh();


			this.uc_scenario_fvs_prepost_variables_effective1.loadvalues();
			p_frmTherm.progressBar1.Value=4;

			

			p_frmTherm.lblMsg.Text = "Rule Definitions: Owner Group Data";
			p_frmTherm.lblMsg.Refresh();

			this.uc_scenario_owner_groups1.loadvalues();
			p_frmTherm.progressBar1.Value=5;
			p_frmTherm.lblMsg.Text = "Rule Definitions: Cost And Revenue Data";
			p_frmTherm.lblMsg.Refresh();
			this.uc_scenario_costs1.loadvalues();
			p_frmTherm.progressBar1.Value=6;
			p_frmTherm.lblMsg.Text = "Rule Definitions: Plot Filter Data";
			p_frmTherm.lblMsg.Refresh();
			this.uc_scenario_filter1.loadvalues();
			this.uc_scenario_cond_filter1.loadvalues();
			this.uc_scenario_psite1.loadvalues();
			p_frmTherm.progressBar1.Value=7;
			p_frmTherm.lblMsg.Text = "Rule Definitions: Wood Processing Site Data";
			p_frmTherm.lblMsg.Refresh();
			this.uc_scenario_filter1.loadvalues();
			p_frmTherm.progressBar1.Value=8;
			this.uc_scenario_run1.chkTreeSumTable();             //make sure table has records
			this.uc_scenario_run1.chkPlotTableForTravelTimes();  //make sure table has travel times
			p_frmTherm.Close();
			p_frmTherm = null;
			this.m_lrulesfirsttime=false;

		}
		public void SaveRuleDefinitions()
		{
			int savestatus;

			frmMain.g_sbpInfo.Text = "Saving scenario rule definitions...Stand by";
			if (m_lrulesfirsttime==false)
			{
				//savestatus=this.uc_scenario_treatment_intensity1.savevalues();
				//savestatus=this.uc_scenario_ffe1.savevalues();
				savestatus=this.uc_scenario_fvs_prepost_variables_effective1.savevalues();
				savestatus=this.uc_scenario_fvs_prepost_optimization1.savevalues();
				savestatus=this.uc_scenario_fvs_prepost_variables_tiebreaker1.savevalues();
				savestatus=this.uc_scenario_owner_groups1.savevalues();
				savestatus=this.uc_scenario_costs1.savevalues();
				savestatus=this.uc_scenario_psite1.savevalues();
				savestatus=this.uc_scenario_filter1.savevalues();
				savestatus=this.uc_scenario_cond_filter1.savevalues();
			}
			this.uc_scenario_notes1.SaveScenarioNotes();
			this.uc_scenario1.UpdateDescription();
			this.m_bSave=false;
			frmMain.g_sbpInfo.Text = "Ready";

		}
		

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			CheckToSave();
			this.Close();
		}
		private void CheckToSave()
		{
			if (m_bSave)
			{
				DialogResult result = MessageBox.Show("Save Changes Y/N","Scenario",System.Windows.Forms.MessageBoxButtons.YesNo,System.Windows.Forms.MessageBoxIcon.Question);
				if (result == System.Windows.Forms.DialogResult.Yes)
				{
					this.SaveRuleDefinitions();
				}
				
			}
		}

		private void vScrollBar_Scroll(Object sender, ScrollEventArgs e)
		{
		}

		private void vScrollBar_ValueChanged(Object sender, EventArgs e)
		{
			this.RulesRepositionControls();
		}
		private void vScrollBar_MouseWheel(Object sender, MouseEventArgs e)
		{
		}
		private void frmScenario_MouseWheel(Object sender, MouseEventArgs e)
		{
			if (e.Delta == 120)
			{
				if (this.m_vScrollBar.Value > this.m_vScrollBar.Minimum) 
				{
					if (this.m_vScrollBar.Value - this.m_vScrollBar.LargeChange < this.m_vScrollBar.Minimum) 
					{
						this.m_vScrollBar.Value = this.m_vScrollBar.Minimum;
					}
					else 
					{
						this.m_vScrollBar.Value = this.m_vScrollBar.Value - 
							this.m_vScrollBar.LargeChange;
					}
					this.RulesRepositionControls();
				}
			}
			else 
			{
				if (this.m_vScrollBar.Value <= this.m_vScrollBar.Maximum) 
				{
					if (this.m_vScrollBar.Value + this.m_vScrollBar.LargeChange > 
						this.m_vScrollBar.Maximum) 
					{
						this.m_vScrollBar.Value = this.m_vScrollBar.Maximum;
					}
					else 
					{
						this.m_vScrollBar.Value = this.m_vScrollBar.Value +
							this.m_vScrollBar.LargeChange;
					}
					this.RulesRepositionControls();
				}

			}
		}
		public void HScrollRepositionControls()
		{
			try
			{
				int intScroll;
				double dblVal = Convert.ToDouble(this.m_hScrollBar.Value);
				double dblMax = Convert.ToDouble(this.m_hScrollBar.Maximum);

				if (this.m_hScrollBar.Value == 0)
				{
					this.m_dblHScrollNewPerc = 0;
				}
				else
				{
					this.m_dblHScrollNewPerc = (dblVal / dblMax);
				}

				if (this.m_dblHScrollNewPerc == 0 && this.m_dblHScrollOldPerc  > 0)
				{
					intScroll =  (int)(this.m_intHScrollMaxSize * this.m_dblHScrollOldPerc);

				}
				else if (this.m_dblHScrollNewPerc > 0 && this.m_dblHScrollOldPerc == 0)
				{
					intScroll = (int)(this.m_intHScrollMaxSize * this.m_dblHScrollNewPerc);
					intScroll = -1 * intScroll;
				}
				else if (this.m_dblHScrollNewPerc > 0 && this.m_dblHScrollOldPerc > 0)
				{
				
					intScroll = (int)(this.m_intHScrollMaxSize * this.m_dblHScrollNewPerc) - (int)(this.m_intHScrollMaxSize * this.m_dblHScrollOldPerc);
					intScroll = -1 * intScroll;
			
				}
				else
				{
					intScroll = 0;
				}
				for (int z=0;z<=this.Controls.Count-1;z++)
				{
					this.Controls[z].Left = this.Controls[z].Left + intScroll;
				}
				this.m_dblHScrollOldPerc = this.m_dblHScrollNewPerc;
			}
			catch
			{
			}
		}
		
		private void button1_Click(object sender, System.EventArgs e)
		{
			MessageBox.Show("thisHeight=" + this.Height.ToString() + " this.m_vScrollBar.Value=" + this.m_vScrollBar.Value.ToString() + " this.m_vScrollBar.Maximum=" + this.m_vScrollBar.Maximum.ToString());
		}
		private void grpboxMenu_Paint(object sender, PaintEventArgs e) 
		{
		}

		private void btnRun_Click(object sender, System.EventArgs e)
		{
			int intAvailWd = this.ParentForm.ClientSize.Width - ((frmMain)this.ParentForm).grpboxLeft.Left - ((frmMain)this.ParentForm).grpboxLeft.Width - 20;
			int intAvailHt = this.ParentForm.ClientSize.Height - ((frmMain)this.ParentForm).tlbMain.Top - ((frmMain)this.ParentForm).tlbMain.Height - 20;
			utils p_oUtils = new utils();
			if (p_oUtils.FindWindowLike((IntPtr)((frmMain)this.ParentForm).Handle, "Core Analysis: Run Scenario (" +  this.uc_scenario1.txtScenarioId.Text.Trim() + ")","*",true,true) > 0)
			{
				this.frmRunCoreScenario1.WindowState = System.Windows.Forms.FormWindowState.Normal;
				this.frmRunCoreScenario1.Focus();
				this.frmRunCoreScenario1.Height = intAvailHt; 
				this.frmRunCoreScenario1.Width = intAvailWd; 
				this.frmRunCoreScenario1.Left = 0;
				this.frmRunCoreScenario1.Top = 0;
				return;

			}
			this.frmRunCoreScenario1 = new frmRunCoreScenario(this);
			this.frmRunCoreScenario1.Text = "Core Analysis: Run Scenario (" + this.uc_scenario1.txtScenarioId.Text.Trim() + ")";
			this.frmRunCoreScenario1.MdiParent = this.ParentForm;
			this.frmRunCoreScenario1.btnCancel.Text = "Start";
			this.frmRunCoreScenario1.WindowState = System.Windows.Forms.FormWindowState.Normal;
			this.frmRunCoreScenario1.Enabled=true;
			this.frmRunCoreScenario1.Show();
			
			
			this.frmRunCoreScenario1.Height = intAvailHt; 
			this.frmRunCoreScenario1.Width = intAvailWd; 
			this.frmRunCoreScenario1.Left = 0;
			this.frmRunCoreScenario1.Top = 0;
		}
		private void hScrollBar_ValueChanged(Object sender, EventArgs e)
		{
			try
			{
				int myValue = ((HScrollBar)sender).Value;
				this.HScrollRepositionControls();
				this.m_intHScrollValue = myValue;
			}
			catch
			{
			}

		}

		private void uc_scenario_notes1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void uc_scenario_notes2_Load(object sender, System.EventArgs e)
		{
		
		}

		private void tabControlRules_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			switch (tabControlRules.SelectedTab.Text.Trim().ToUpper())
			{
				case "WOOD PROCESSING SITES":
					if (((Control)this.tbPSites).Enabled)
					    this.uc_scenario_psite1.lblTitle.Text = "Wood Processing Sites";
					else
						this.uc_scenario_psite1.lblTitle.Text = "Wood Processing Sites (Read Only)";

					break;
				case "COST AND REVENUE":
					if (((Control)this.tbCost).Enabled)
						this.uc_scenario_costs1.lblTitle.Text = "Cost And Revenue";
					else
						this.uc_scenario_costs1.lblTitle.Text = "Cost And Revenue (Read Only)";
					break;
				case "FILTER PLOT RECORDS":
					if (((Control)this.tbFilterPlots).Enabled)
						this.uc_scenario_filter1.lblTitle.Text = "Plot Filter";
					else
						this.uc_scenario_filter1.lblTitle.Text = "Plot Filter (Read Only)";
					break;
				case "LAND OWNERSHIP GROUPS":
					if (((Control)this.tbOwners).Enabled)
						this.uc_scenario_owner_groups1.lblTitle.Text = "Owner Groups";
					else
						this.uc_scenario_owner_groups1.lblTitle.Text = "Owner Groups (Read Only)";
					break;
                //case "TREATMENT INTENSITY":
				//	if (((Control)this.tbTreatmentIntensity).Enabled)
				//		this.uc_scenario_treatment_intensity1.lblTitle.Text = "Treatment Intensity";
				//	else
				//		this.uc_scenario_treatment_intensity1.lblTitle.Text = "Treatment Intensity (Read Only)";
				//	break;
				case "RUN":
					if (((Control)this.tbRun).Enabled)
						this.uc_scenario_run1.lblTitle.Text = "Run";
					else
						this.uc_scenario_run1.lblTitle.Text = "Run (Read Only)";
					break;



			}
		}

		private void tlbScenario_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			switch (e.Button.ImageIndex)
			{
					//
					//open scenario
					//
				case 0:
					frmMain.g_oFrmMain.OpenCoreScenario("Open");
					break;
					//
					//new scenario
					//
				case 1:
					frmMain.g_oFrmMain.OpenCoreScenario("New");
					break;
				case 2:
					this.SaveRuleDefinitions();
					break;
				case 3:
					if (this.uc_scenario_open1 != null)
					{
						frmMain.g_oFrmMain.DeleteScenario("core",uc_scenario_open1.txtScenarioId.Text.Trim());
						uc_scenario_open1.lstScenario.Items.Remove(uc_scenario_open1.lstScenario.SelectedItems[0]);
					}
					else
					{
						if (frmMain.g_oFrmMain.DeleteScenario("core",uc_scenario1.txtScenarioId.Text.Trim()))
							this.Close();
					}
					break;
			}
		}

		private void tbRules_Click(object sender, System.EventArgs e)
		{
		}

		private void tabControlScenario_TabIndexChanged(object sender, System.EventArgs e)
		{
			MessageBox.Show("tabControlScenario_TabIndexChanged");
		}

		private void tabControlScenario_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.tabControlRules.Enabled)
			{
				if (tabControlScenario.SelectedTab.Text.Trim().ToUpper()=="RULE DEFINITIONS")
				{
					if (m_lrulesfirsttime==true)
						LoadRuleDefinitions();

				}
				else if (tabControlScenario.SelectedTab.Text.Trim().ToUpper() == "NOTES")
				{
					if (((Control)this.tbNotes).Enabled)
					{
					  this.uc_scenario_notes1.lblTitle.Text = "Notes";
					}
					else
					{
						this.uc_scenario_notes1.lblTitle.Text = "Notes (Read Only)";
					}
				}
				else if (tabControlScenario.SelectedTab.Text.Trim().ToUpper() == "DESCRIPTION")
				{
					if (((Control)this.tbDesc).Enabled)
					{
						this.uc_scenario1.lblTitle.Text = "Description";
					}
					else
					{
						this.uc_scenario1.lblTitle.Text = "Description (Read Only)";
					}
				}
				else if (tabControlScenario.SelectedTab.Text.Trim().ToUpper() == "DATA SOURCES")
				{
					if (((Control)this.tbDataSources).Enabled)
					{
						this.uc_datasource1.lblTitle.Text = "Scenario Datasource";
					}
					else
					{
						this.uc_datasource1.lblTitle.Text = "Scenario Datasource (Read Only)";
					}
				}


			}

			

		}

		private void tbPSites_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void tbPSites_Click(object sender, System.EventArgs e)
		{
		
		}

		private void tbOwners_Click(object sender, System.EventArgs e)
		{
		
		}

		private void tbFVSVariablesSelect_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			try
			{
				if (tabControlFVSVariables.Enabled)
				{
					if (tabControlFVSVariables.SelectedTab.Text.Trim().ToUpper()=="OPTIMIZATION")
					{
						
							if (((Control)this.tbOptimization).Enabled)
							{
						
								uc_scenario_fvs_prepost_optimization1.lblTitle.Text = "Optimization Settings";
							}
							else
								uc_scenario_fvs_prepost_optimization1.lblTitle.Text = "Optimization Settings (Read Only)";
						
					}
					else if (tabControlFVSVariables.SelectedTab.Text.Trim().ToUpper()=="EFFECTIVE")
					{
						
						if (((Control)this.tbEffective).Enabled)
						{
						    uc_scenario_fvs_prepost_variables_effective1.lblTitle.Text  = "Effective Settings";
						}
						else
							uc_scenario_fvs_prepost_variables_effective1.lblTitle.Text = "Effective Settings (Read Only)";
						
					}
					else if (tabControlFVSVariables.SelectedTab.Text.Trim().ToUpper()=="TIE BREAKER")
					{
						
						if (((Control)this.tbTieBreaker).Enabled)
						{
							uc_scenario_fvs_prepost_variables_tiebreaker1.lblTitle.Text  = "Tie Breaker Settings";
						}
						else
							uc_scenario_fvs_prepost_variables_tiebreaker1.lblTitle.Text = "Tie Breaker Settings (Read Only)";
						
					}
					
				}
				
			}
			catch (Exception err)
			{
				MessageBox.Show(err.Message);
			}
		}

		private void tabControlScenario_DrawItem(object sender, System.Windows.Forms.DrawItemEventArgs e)
		{
			TabControl_DrawItem(sender,e,Color.Green,System.Drawing.Brushes.White);
		}
		public void TabControl_DrawItem(object sender, System.Windows.Forms.DrawItemEventArgs e,System.Drawing.Color p_oSelectedBackgroundColor, System.Drawing.Brush p_oSelectedForegroundColor)
		{
			try
			{
				//This line of code will help you to change the apperance like size,name,style.
				System.Windows.Forms.TabControl tc = (System.Windows.Forms.TabControl)sender;
				
				Font f;
				//For background color
				Brush backBrush;
				//For forground color
				Brush foreBrush;
			
				//This construct will hell you to deside which tab page have current focus
				//to change the style.
				if(e.Index == tc.SelectedIndex) //tabControlScenario.SelectedIndex)
				{
					//This line of code will help you to change the apperance like size,name,style.
					f = new Font(e.Font, FontStyle.Regular | FontStyle.Regular);
					f = new Font(e.Font,FontStyle.Regular);

					backBrush = new System.Drawing.SolidBrush(p_oSelectedBackgroundColor);
					//foreBrush = Brushes.White;
					foreBrush = p_oSelectedForegroundColor;
				}
				else
				{
					f = e.Font;
					backBrush = new SolidBrush(SystemColors.Control); 
					foreBrush = new SolidBrush(e.ForeColor);
				}
                
				//To set the alignment of the caption.
				string tabName = tc.TabPages[e.Index].Text; //tc.Indexe.IndextabControlScenario.TabPages[e.Index].Text;
				StringFormat sf = new StringFormat();
				sf.Alignment = StringAlignment.Center;
	
				//This will help you to fill the interior portion of
				//selected tabpage.
				e.Graphics.FillRectangle(backBrush, e.Bounds);
				Rectangle r = e.Bounds;
				r = new Rectangle(r.X, r.Y + 3, r.Width, r.Height - 3);
				//r = new Rectangle(r.X, r.Y + 4 , r.Width, r.Height - 4 );
				e.Graphics.DrawString(tabName, f, foreBrush, r, sf);

				sf.Dispose();
				if(e.Index == tc.SelectedIndex)
				{
					f.Dispose();
					backBrush.Dispose();
				}
				else
				{
					backBrush.Dispose();
					foreBrush.Dispose();
				}
			}
			catch(Exception Ex)
			{
				MessageBox.Show(Ex.Message.ToString(),"Error Occured",MessageBoxButtons.OK,MessageBoxIcon.Information);
					
			}

		}

		private void tabControlRules_DrawItem(object sender, System.Windows.Forms.DrawItemEventArgs e)
		{
			TabControl_DrawItem(sender,e,Color.DarkBlue,System.Drawing.Brushes.White);
		}

		private void tbFVSVariablesSelect_DrawItem(object sender, System.Windows.Forms.DrawItemEventArgs e)
		{
			TabControl_DrawItem(sender,e,Color.DarkRed,System.Drawing.Brushes.White);
		}

		private void tabControlRules_TabIndexChanged(object sender, System.EventArgs e)
		{
			MessageBox.Show("tabControlRules_TabIndexChanged");
		}

		private void tabControlScenario_Leave(object sender, System.EventArgs e)
		{
			
		}

		private void tabControlScenario_Enter(object sender, System.EventArgs e)
		{
			
		}

		private void tabControlScenario_ChangeUICues(object sender, System.Windows.Forms.UICuesEventArgs e)
		{
		
		}
		public void EnableTabPage(System.Windows.Forms.TabPage p_oTabPage,bool p_bEnable)
		{
				((Control)p_oTabPage).Enabled = p_bEnable;	
		}

		public void EnableTabPage(System.Windows.Forms.TabControl p_oTabControl,string p_strTabPageNameList,bool p_bEnable)
		{
			int y,z;
			string[] strTabPageNameArray = frmMain.g_oUtils.ConvertListToArray(p_strTabPageNameList,",");
			for (y=0;y<=p_oTabControl.TabPages.Count-1;y++)
			{
				for (z=0;z<=strTabPageNameArray.Length-1;z++)
				{
					if (p_oTabControl.TabPages[y].Name.Trim().ToUpper()==strTabPageNameArray[z].Trim().ToUpper())
					{
						EnableTabPage(p_oTabControl.TabPages[y],p_bEnable);
					}
				}
			}
		}

		public void EnableControl(string p_strControlNameList,bool p_bEnable)
		{
			int x,y;
		    string[] strControlNameArray = frmMain.g_oUtils.ConvertListToArray(p_strControlNameList,",");

			foreach (System.Windows.Forms.Control c in this.Controls)
			{
				if (c.Name.Trim().Length > 0)
				{
					for (x=0;x<=strControlNameArray.Length -1;x++)
					{
						if (c.Name.Trim().ToUpper()==strControlNameArray[x].Trim().ToUpper())
						{
							c.Enabled=p_bEnable;
							break;
						}
					}
				}
			}
		}

		private void frmScenario_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (frmMain.g_oDelegate.CurrentThreadProcessIdle==false)
			{
				e.Cancel = true;
				return;
			}
			CheckToSave();
		}

		


		
		
		

	}
}
