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
	public class frmCoreScenario : System.Windows.Forms.Form
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
		public FIA_Biosum_Manager.uc_scenario_costs uc_scenario_costs1;
		public FIA_Biosum_Manager.uc_select_list_item uc_select_list_item1;
		public FIA_Biosum_Manager.uc_scenario_filter uc_scenario_filter1;
		public FIA_Biosum_Manager.uc_scenario_filter uc_scenario_cond_filter1;
		public FIA_Biosum_Manager.uc_scenario_psite uc_scenario_psite1;
		public FIA_Biosum_Manager.uc_scenario_open uc_scenario_open1;
		public FIA_Biosum_Manager.uc_core_scenario_run uc_scenario_run1=null;
		public FIA_Biosum_Manager.uc_scenario_notes uc_scenario_notes1;
        public FIA_Biosum_Manager.uc_scenario_processor_scenario_select uc_scenario_processor_scenario_select1;
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
        public bool m_bSave = false;
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
        private TabPage tbCostsAndRevenue;
        public TabControl tabControlCosts;
        private TabPage tbCosts;
        private TabPage tbProcessorScenario;

        public const int TOTALCYCLES = 4;
		
		
		public bool m_bPopup = false;

        public FIA_Biosum_Manager.CoreAnalysisScenarioItem m_oCoreAnalysisScenarioItem = new CoreAnalysisScenarioItem();
        public FIA_Biosum_Manager.CoreAnalysisScenarioItem_Collection m_oCoreAnalysisScenarioItem_Collection = new CoreAnalysisScenarioItem_Collection();
        public FIA_Biosum_Manager.CoreAnalysisScenarioTools m_oCoreAnalysisScenarioTools = new CoreAnalysisScenarioTools();
        private ToolBarButton btnScenarioProperties;
        private ToolBarButton tlbSeparator;
        private ToolBarButton btnScenarioCopy;
        public FIA_Biosum_Manager.CoreAnalysisScenarioItem m_oSavCoreAnalysisScenarioItem = new CoreAnalysisScenarioItem();
        

		

		public frmCoreScenario(FIA_Biosum_Manager.frmMain p_frmMain)
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
				uc_scenario1.ReferenceCoreAnalysisScenarioForm=this;
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
				this.tbCosts.Controls.Add(uc_scenario_costs1);
              
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
                //
                //processor scenario select
                //
                this.uc_scenario_processor_scenario_select1 = new uc_scenario_processor_scenario_select();
                this.tbProcessorScenario.Controls.Add(this.uc_scenario_processor_scenario_select1);
                this.uc_scenario_processor_scenario_select1.Dock = System.Windows.Forms.DockStyle.Fill;
                this.uc_scenario_processor_scenario_select1.ReferenceCoreScenarioForm = this;


				//rule execute run
				this.uc_scenario_run1 = new uc_core_scenario_run();
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

		public frmCoreScenario()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmCoreScenario));
            this.tlbScenario = new System.Windows.Forms.ToolBar();
            this.btnScenarioNew = new System.Windows.Forms.ToolBarButton();
            this.btnScenarioOpen = new System.Windows.Forms.ToolBarButton();
            this.toolBarButton1 = new System.Windows.Forms.ToolBarButton();
            this.btnScenarioSave = new System.Windows.Forms.ToolBarButton();
            this.btnScenarioDelete = new System.Windows.Forms.ToolBarButton();
            this.btnScenarioProperties = new System.Windows.Forms.ToolBarButton();
            this.tlbSeparator = new System.Windows.Forms.ToolBarButton();
            this.btnScenarioCopy = new System.Windows.Forms.ToolBarButton();
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
            this.tbCostsAndRevenue = new System.Windows.Forms.TabPage();
            this.tabControlCosts = new System.Windows.Forms.TabControl();
            this.tbCosts = new System.Windows.Forms.TabPage();
            this.tbProcessorScenario = new System.Windows.Forms.TabPage();
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
            this.tbCostsAndRevenue.SuspendLayout();
            this.tabControlCosts.SuspendLayout();
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
            this.btnScenarioDelete,
            this.btnScenarioProperties,
            this.tlbSeparator,
            this.btnScenarioCopy});
            this.tlbScenario.DropDownArrows = true;
            this.tlbScenario.ImageList = this.imageList1;
            this.tlbScenario.Location = new System.Drawing.Point(0, 0);
            this.tlbScenario.Name = "tlbScenario";
            this.tlbScenario.ShowToolTips = true;
            this.tlbScenario.Size = new System.Drawing.Size(663, 44);
            this.tlbScenario.TabIndex = 42;
            this.tlbScenario.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.tlbScenario_ButtonClick);
            // 
            // btnScenarioNew
            // 
            this.btnScenarioNew.ImageIndex = 1;
            this.btnScenarioNew.Name = "btnScenarioNew";
            this.btnScenarioNew.Text = "New";
            this.btnScenarioNew.ToolTipText = "New";
            // 
            // btnScenarioOpen
            // 
            this.btnScenarioOpen.ImageIndex = 0;
            this.btnScenarioOpen.Name = "btnScenarioOpen";
            this.btnScenarioOpen.Text = "Open";
            this.btnScenarioOpen.ToolTipText = "Open";
            // 
            // toolBarButton1
            // 
            this.toolBarButton1.Name = "toolBarButton1";
            this.toolBarButton1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
            // 
            // btnScenarioSave
            // 
            this.btnScenarioSave.ImageIndex = 2;
            this.btnScenarioSave.Name = "btnScenarioSave";
            this.btnScenarioSave.Text = "Save";
            this.btnScenarioSave.ToolTipText = "Save";
            // 
            // btnScenarioDelete
            // 
            this.btnScenarioDelete.ImageIndex = 3;
            this.btnScenarioDelete.Name = "btnScenarioDelete";
            this.btnScenarioDelete.Text = "Delete";
            this.btnScenarioDelete.ToolTipText = "Delete";
            // 
            // btnScenarioProperties
            // 
            this.btnScenarioProperties.ImageIndex = 4;
            this.btnScenarioProperties.Name = "btnScenarioProperties";
            this.btnScenarioProperties.Text = "Properties";
            this.btnScenarioProperties.Visible = false;
            // 
            // tlbSeparator
            // 
            this.tlbSeparator.Name = "tlbSeparator";
            this.tlbSeparator.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
            // 
            // btnScenarioCopy
            // 
            this.btnScenarioCopy.ImageIndex = 5;
            this.btnScenarioCopy.Name = "btnScenarioCopy";
            this.btnScenarioCopy.Text = "Copy";
            this.btnScenarioCopy.ToolTipText = "Copy scenario values to this scenario";
            this.btnScenarioCopy.Visible = false;
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "open-file-icon.png");
            this.imageList1.Images.SetKeyName(1, "");
            this.imageList1.Images.SetKeyName(2, "");
            this.imageList1.Images.SetKeyName(3, "deletered.png");
            this.imageList1.Images.SetKeyName(4, "properties.png");
            this.imageList1.Images.SetKeyName(5, "copy.bmp");
            // 
            // txtDropDown
            // 
            this.txtDropDown.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.txtDropDown.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDropDown.ForeColor = System.Drawing.SystemColors.ControlText;
            this.txtDropDown.HideSelection = false;
            this.txtDropDown.Location = new System.Drawing.Point(224, 416);
            this.txtDropDown.Multiline = true;
            this.txtDropDown.Name = "txtDropDown";
            this.txtDropDown.ReadOnly = true;
            this.txtDropDown.Size = new System.Drawing.Size(24, 24);
            this.txtDropDown.TabIndex = 12;
            this.txtDropDown.Visible = false;
            // 
            // imgSize
            // 
            this.imgSize.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imgSize.ImageStream")));
            this.imgSize.TransparentColor = System.Drawing.Color.Transparent;
            this.imgSize.Images.SetKeyName(0, "");
            this.imgSize.Images.SetKeyName(1, "");
            // 
            // imageList2
            // 
            this.imageList2.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList2.ImageStream")));
            this.imageList2.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList2.Images.SetKeyName(0, "");
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
            this.tabControlScenario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControlScenario.ItemSize = new System.Drawing.Size(100, 18);
            this.tabControlScenario.Location = new System.Drawing.Point(0, 32);
            this.tabControlScenario.Name = "tabControlScenario";
            this.tabControlScenario.SelectedIndex = 0;
            this.tabControlScenario.Size = new System.Drawing.Size(648, 360);
            this.tabControlScenario.TabIndex = 41;
            this.tabControlScenario.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.tabControlScenario_DrawItem);
            this.tabControlScenario.SelectedIndexChanged += new System.EventHandler(this.tabControlScenario_SelectedIndexChanged);
            this.tabControlScenario.TabIndexChanged += new System.EventHandler(this.tabControlScenario_TabIndexChanged);
            this.tabControlScenario.Enter += new System.EventHandler(this.tabControlScenario_Enter);
            this.tabControlScenario.Leave += new System.EventHandler(this.tabControlScenario_Leave);
            this.tabControlScenario.ChangeUICues += new System.Windows.Forms.UICuesEventHandler(this.tabControlScenario_ChangeUICues);
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
            this.tabControlRules.Controls.Add(this.tbCostsAndRevenue);
            this.tabControlRules.Controls.Add(this.tbPSites);
            this.tabControlRules.Controls.Add(this.tbFilterPlots);
            this.tabControlRules.Controls.Add(this.tbFilterCond);
            this.tabControlRules.Controls.Add(this.tbFVSVariables);
            this.tabControlRules.Controls.Add(this.tbRun);
            this.tabControlRules.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlRules.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.tabControlRules.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControlRules.Location = new System.Drawing.Point(0, 0);
            this.tabControlRules.Name = "tabControlRules";
            this.tabControlRules.SelectedIndex = 0;
            this.tabControlRules.Size = new System.Drawing.Size(640, 334);
            this.tabControlRules.TabIndex = 0;
            this.tabControlRules.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.tabControlRules_DrawItem);
            this.tabControlRules.SelectedIndexChanged += new System.EventHandler(this.tabControlRules_SelectedIndexChanged);
            this.tabControlRules.TabIndexChanged += new System.EventHandler(this.tabControlRules_TabIndexChanged);
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
            // tbCostsAndRevenue
            // 
            this.tbCostsAndRevenue.Controls.Add(this.tabControlCosts);
            this.tbCostsAndRevenue.Location = new System.Drawing.Point(4, 22);
            this.tbCostsAndRevenue.Name = "tbCostsAndRevenue";
            this.tbCostsAndRevenue.Size = new System.Drawing.Size(632, 308);
            this.tbCostsAndRevenue.TabIndex = 10;
            this.tbCostsAndRevenue.Text = "Cost And Revenue";
            this.tbCostsAndRevenue.UseVisualStyleBackColor = true;
            // 
            // tabControlCosts
            // 
            this.tabControlCosts.Controls.Add(this.tbCosts);
            this.tabControlCosts.Controls.Add(this.tbProcessorScenario);
            this.tabControlCosts.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlCosts.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.tabControlCosts.Location = new System.Drawing.Point(0, 0);
            this.tabControlCosts.Name = "tabControlCosts";
            this.tabControlCosts.SelectedIndex = 0;
            this.tabControlCosts.Size = new System.Drawing.Size(632, 308);
            this.tabControlCosts.TabIndex = 1;
            this.tabControlCosts.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.tabControlCosts_DrawItem);
            this.tabControlCosts.SelectedIndexChanged += new System.EventHandler(this.tabControlCosts_SelectedIndexChanged);
            // 
            // tbCosts
            // 
            this.tbCosts.Location = new System.Drawing.Point(4, 22);
            this.tbCosts.Name = "tbCosts";
            this.tbCosts.Padding = new System.Windows.Forms.Padding(3);
            this.tbCosts.Size = new System.Drawing.Size(624, 282);
            this.tbCosts.TabIndex = 0;
            this.tbCosts.Text = "Haul Costs";
            this.tbCosts.UseVisualStyleBackColor = true;
            // 
            // tbProcessorScenario
            // 
            this.tbProcessorScenario.Location = new System.Drawing.Point(4, 22);
            this.tbProcessorScenario.Name = "tbProcessorScenario";
            this.tbProcessorScenario.Padding = new System.Windows.Forms.Padding(3);
            this.tbProcessorScenario.Size = new System.Drawing.Size(624, 282);
            this.tbProcessorScenario.TabIndex = 1;
            this.tbProcessorScenario.Text = "Processor Scenario";
            this.tbProcessorScenario.UseVisualStyleBackColor = true;
            // 
            // tbPSites
            // 
            this.tbPSites.Location = new System.Drawing.Point(4, 22);
            this.tbPSites.Name = "tbPSites";
            this.tbPSites.Size = new System.Drawing.Size(632, 308);
            this.tbPSites.TabIndex = 4;
            this.tbPSites.Text = "Wood Processing Sites";
            this.tbPSites.Click += new System.EventHandler(this.tbPSites_Click);
            this.tbPSites.Enter += new System.EventHandler(this.tbPSites_Enter);
            // 
            // tbFilterPlots
            // 
            this.tbFilterPlots.Location = new System.Drawing.Point(4, 22);
            this.tbFilterPlots.Name = "tbFilterPlots";
            this.tbFilterPlots.Size = new System.Drawing.Size(632, 308);
            this.tbFilterPlots.TabIndex = 5;
            this.tbFilterPlots.Text = "Filter Plot Records";
            // 
            // tbFilterCond
            // 
            this.tbFilterCond.Location = new System.Drawing.Point(4, 22);
            this.tbFilterCond.Name = "tbFilterCond";
            this.tbFilterCond.Size = new System.Drawing.Size(632, 308);
            this.tbFilterCond.TabIndex = 9;
            this.tbFilterCond.Text = "Filter Condition Records";
            // 
            // tbFVSVariables
            // 
            this.tbFVSVariables.Controls.Add(this.tabControlFVSVariables);
            this.tbFVSVariables.Location = new System.Drawing.Point(4, 22);
            this.tbFVSVariables.Name = "tbFVSVariables";
            this.tbFVSVariables.Size = new System.Drawing.Size(632, 308);
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
            this.tabControlFVSVariables.Size = new System.Drawing.Size(632, 308);
            this.tabControlFVSVariables.TabIndex = 0;
            this.tabControlFVSVariables.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.tbFVSVariablesSelect_DrawItem);
            this.tabControlFVSVariables.SelectedIndexChanged += new System.EventHandler(this.tbFVSVariablesSelect_SelectedIndexChanged);
            // 
            // tbEffective
            // 
            this.tbEffective.Location = new System.Drawing.Point(4, 22);
            this.tbEffective.Name = "tbEffective";
            this.tbEffective.Size = new System.Drawing.Size(624, 282);
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
            this.tbRun.Location = new System.Drawing.Point(4, 22);
            this.tbRun.Name = "tbRun";
            this.tbRun.Size = new System.Drawing.Size(632, 308);
            this.tbRun.TabIndex = 6;
            this.tbRun.Text = "Run";
            // 
            // frmCoreScenario
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(663, 454);
            this.Controls.Add(this.tlbScenario);
            this.Controls.Add(this.tabControlScenario);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.txtDropDown);
            this.Name = "frmCoreScenario";
            this.Text = "Case Study Scenario";
            this.Activated += new System.EventHandler(this.frmCoreScenario_Activated);
            this.Closing += new System.ComponentModel.CancelEventHandler(this.frmScenario_Closing);
            this.Load += new System.EventHandler(this.frmscenarioScenario_Load);
            this.Click += new System.EventHandler(this.frmScenario_Click);
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.frmScenario_MouseDown);
            this.Resize += new System.EventHandler(this.frmScenario_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dataView1)).EndInit();
            this.tabControlScenario.ResumeLayout(false);
            this.tbRules.ResumeLayout(false);
            this.tabControlRules.ResumeLayout(false);
            this.tbCostsAndRevenue.ResumeLayout(false);
            this.tabControlCosts.ResumeLayout(false);
            this.tbFVSVariables.ResumeLayout(false);
            this.tabControlFVSVariables.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

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
				frmCoreScenario frmTemp = new frmCoreScenario((frmMain)this.ParentForm);
				frmTemp.Visible=false;
				frmTemp.uc_datasource1.Visible = false;
				frmTemp.uc_scenario_notes1.Visible=false;
				frmTemp.uc_scenario1.Visible=false;
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

			this.uc_scenario1.ReferenceCoreAnalysisScenarioForm=this;

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
				frmCoreScenario frmTemp = new frmCoreScenario((frmMain)this.ParentForm);
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
            int x;
			frmTherm p_frmTherm = new frmTherm();
			p_frmTherm.lblMsg.Text = "";
			p_frmTherm.progressBar1.Minimum = 0;
			p_frmTherm.progressBar1.Maximum = 8;
			p_frmTherm.btnCancel.Visible=false;
			p_frmTherm.lblMsg.Visible=true;
			p_frmTherm.Show();
			p_frmTherm.progressBar1.Value=1;
			p_frmTherm.progressBar1.Value=2;
            Queries oQueries = new Queries();
            
            this.m_oCoreAnalysisScenarioItem.ScenarioId = this.uc_scenario1.txtScenarioId.Text.Trim();

            this.m_oCoreAnalysisScenarioTools.LoadAll(
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\core\\db\\scenario_core_rule_definitions.mdb",
                oQueries, m_oCoreAnalysisScenarioItem.ScenarioId,
                m_oCoreAnalysisScenarioItem_Collection);
            //find the current scenario
            for (x = 0; x <= m_oCoreAnalysisScenarioItem_Collection.Count - 1; x++)
            {
                if (m_oCoreAnalysisScenarioItem_Collection.Item(x).ScenarioId.Trim().ToUpper() ==
                    m_oCoreAnalysisScenarioItem.ScenarioId.Trim().ToUpper())
                    break;
            }

            this.m_oCoreAnalysisScenarioItem.Copy(m_oCoreAnalysisScenarioItem_Collection.Item(x), m_oSavCoreAnalysisScenarioItem);
            this.m_oCoreAnalysisScenarioItem = m_oCoreAnalysisScenarioItem_Collection.Item(x);

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
            this.uc_scenario_processor_scenario_select1.loadvalues();
			p_frmTherm.progressBar1.Value=6;
			p_frmTherm.lblMsg.Text = "Rule Definitions: Plot Filter Data";
			p_frmTherm.lblMsg.Refresh();
			this.uc_scenario_filter1.loadvalues(false);
			this.uc_scenario_cond_filter1.loadvalues(false);
			this.uc_scenario_psite1.loadvalues();
			p_frmTherm.progressBar1.Value=7;
			p_frmTherm.lblMsg.Text = "Rule Definitions: Wood Processing Site Data";
			p_frmTherm.lblMsg.Refresh();
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
            int x;

			frmMain.g_sbpInfo.Text = "Saving scenario rule definitions...Stand by";
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 13;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;

			if (m_lrulesfirsttime==false)
			{
				savestatus=this.uc_scenario_fvs_prepost_variables_effective1.savevalues();
                if (FIA_Biosum_Manager.RunCore.g_bCoreRun) FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
				savestatus=this.uc_scenario_fvs_prepost_optimization1.savevalues();
                if (FIA_Biosum_Manager.RunCore.g_bCoreRun) FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
				savestatus=this.uc_scenario_fvs_prepost_variables_tiebreaker1.savevalues();
                if (FIA_Biosum_Manager.RunCore.g_bCoreRun) FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
				savestatus=this.uc_scenario_owner_groups1.savevalues();
                if (FIA_Biosum_Manager.RunCore.g_bCoreRun) FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
				savestatus=this.uc_scenario_costs1.savevalues();
                if (FIA_Biosum_Manager.RunCore.g_bCoreRun) FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
				savestatus=this.uc_scenario_psite1.savevalues();
                if (FIA_Biosum_Manager.RunCore.g_bCoreRun) FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
				savestatus=this.uc_scenario_filter1.savevalues();
                if (FIA_Biosum_Manager.RunCore.g_bCoreRun) FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
				savestatus=this.uc_scenario_cond_filter1.savevalues();
                if (FIA_Biosum_Manager.RunCore.g_bCoreRun) FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
                this.uc_scenario_processor_scenario_select1.savevalues();
                if (FIA_Biosum_Manager.RunCore.g_bCoreRun) FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
			}
			this.uc_scenario_notes1.SaveScenarioNotes();
            if (FIA_Biosum_Manager.RunCore.g_bCoreRun) FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
			this.uc_scenario1.UpdateDescription();
            if (FIA_Biosum_Manager.RunCore.g_bCoreRun) FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            Queries oQueries = new Queries();
            this.m_oCoreAnalysisScenarioTools.LoadAll(
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\core\\db\\scenario_core_rule_definitions.mdb",
                oQueries, m_oCoreAnalysisScenarioItem.ScenarioId,
                m_oCoreAnalysisScenarioItem_Collection);
            //find the current scenario
            for (x = 0; x <= m_oCoreAnalysisScenarioItem_Collection.Count - 1; x++)
            {
                if (m_oCoreAnalysisScenarioItem_Collection.Item(x).ScenarioId.Trim().ToUpper() ==
                    m_oCoreAnalysisScenarioItem.ScenarioId.Trim().ToUpper())
                    break;
            }
            if (FIA_Biosum_Manager.RunCore.g_bCoreRun) FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            this.m_oCoreAnalysisScenarioItem.Copy(m_oCoreAnalysisScenarioItem_Collection.Item(x), m_oSavCoreAnalysisScenarioItem);
            if (FIA_Biosum_Manager.RunCore.g_bCoreRun) FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
            this.m_oCoreAnalysisScenarioItem = m_oCoreAnalysisScenarioItem_Collection.Item(x);

            
			this.m_bSave=false;
			frmMain.g_sbpInfo.Text = "Ready";

		}
		

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			DialogResult result = CheckToSave();
            switch (result)
            {
                case (DialogResult.Cancel):
                    break;
                default:
                    this.Close();
                    break;
                    
            }
			
		}
		private DialogResult CheckToSave()
		{
            DialogResult result=DialogResult.No;
			if (m_bSave)
			{
				result = MessageBox.Show("Save Changes Y/N","Scenario",System.Windows.Forms.MessageBoxButtons.YesNoCancel,System.Windows.Forms.MessageBoxIcon.Question);
                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    this.SaveRuleDefinitions();
                }
			}
            return result;
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
					if (((Control)this.tbCosts).Enabled)
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
                case 4:
                    if (this.m_lrulesfirsttime == true)
                        LoadRuleDefinitions();
                    else
                    {
                        SaveRuleDefinitions();
                       
                    }
                    FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
                    frmTemp.Text = "FIA Biosum";
                    frmTemp.AutoScroll = false;
                    uc_textbox uc_textbox1 = new uc_textbox();
                    frmTemp.Controls.Add(uc_textbox1);
                    uc_textbox1.Dock = DockStyle.Fill;
                    uc_textbox1.lblTitle.Text = "Properties";
                    uc_textbox1.TextValue = m_oCoreAnalysisScenarioTools.ScenarioProperties(m_oCoreAnalysisScenarioItem);
                    frmTemp.Show();
                    break;
                case 5:
                    CopyScenario();
                    break;
			}
		}
        private void CopyScenario()
        {
            frmDialog frmTemp = new frmDialog();
            frmTemp.Initialize_Scenario_Core_Scenario_Copy();
            frmTemp.Text = "FIA Biosum";
            if (m_oCoreAnalysisScenarioItem.ScenarioId.Trim().Length == 0) LoadRuleDefinitions();

            frmTemp.uc_scenario_core_scenario_copy1.ReferenceCurrentScenarioItem = m_oCoreAnalysisScenarioItem;
            frmTemp.uc_scenario_core_scenario_copy1.loadvalues();
            DialogResult result = frmTemp.ShowDialog(this);
            if (result == DialogResult.OK)
            {

                frmMain.g_sbpInfo.Text = "Copying scenario rule definitions...Stand by";

                this.uc_scenario_fvs_prepost_variables_effective1.loadvalues_FromProperties();
                
                
                this.uc_scenario_owner_groups1.loadvalues();
                
                
                this.uc_scenario_costs1.loadvalues();
                this.uc_scenario_processor_scenario_select1.loadvalues();
                
                
                this.uc_scenario_filter1.loadvalues(true);
                this.uc_scenario_cond_filter1.loadvalues(true);
                this.uc_scenario_psite1.loadvalues_FromProperties();
                
                this.uc_scenario_run1.chkTreeSumTable();             //make sure table has records
                this.uc_scenario_run1.chkPlotTableForTravelTimes();  //make sure table has travel times
                frmMain.g_sbpInfo.Text = "Ready";
                m_bSave = true;
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
			
				//This construct will help you to deside which tab page have current focus
				//to change the style.
				if(e.Index == tc.SelectedIndex) //tabControlScenario.SelectedIndex)
				{
					//This line of code will help you to change the apperance like size,name,style.
					f = new Font(e.Font, FontStyle.Regular | FontStyle.Regular);
					f = new Font(e.Font,FontStyle.Regular);

					backBrush = new System.Drawing.SolidBrush(p_oSelectedBackgroundColor);
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
            if (frmMain.g_oFrmMain.tlbMain.Enabled == false)
            {
                MessageBox.Show("!!Finish Edit Session Before Closing!!", "FIA Biosum");
                e.Cancel = true;
            }
            else
            {
                DialogResult result = CheckToSave();
                switch (result)
                {
                    case (DialogResult.Cancel):
                        e.Cancel = true;
                        break;
                    default:
                        break;

                }
            }
		}

        private void tabControlCosts_DrawItem(object sender, DrawItemEventArgs e)
        {
            TabControl_DrawItem(sender, e, Color.DarkRed, System.Drawing.Brushes.White);
        }

        private void tabControlCosts_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (this.tabControlCosts.Enabled)
                {
                    if (tabControlCosts.SelectedTab.Text.Trim().ToUpper() == "HAUL COSTS")
                    {

                        if (((Control)this.tbCosts).Enabled)
                        {
                            this.uc_scenario_costs1.lblTitle.Text = "Costs And Revenue";
                        }
                        else
                            this.uc_scenario_costs1.lblTitle.Text = "Costs And Revenue (Read Only)";

                    }
                    else if (tabControlFVSVariables.SelectedTab.Text.Trim().ToUpper() == "PROCESSOR SCENARIO")
                    {

                        if (((Control)this.tbProcessorScenario).Enabled)
                        {
                        }
                        else
                        {
                        }

                    }
                    

                }

            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }

        private void frmCoreScenario_Activated(object sender, EventArgs e)
        {
            if (uc_scenario_run1 != null)
            {
                this.uc_scenario_run1.groupBox1_Resize();
            }
        }

		


		
		
		

	}
    public class CoreAnalysisScenarioItem
    {
        
        public EffectiveVariablesItem_Collection m_oEffectiveVariablesItem_Collection = new EffectiveVariablesItem_Collection();
        public OptimizationVariableItem_Collection m_oOptimizationVariableItem_Collection = new OptimizationVariableItem_Collection();
        public TieBreaker_Collection m_oTieBreaker_Collection = new TieBreaker_Collection();
        public RxIntensityItem_Collection m_oRxIntensityItem_Collection = new RxIntensityItem_Collection();
        public ProcessingSiteItem_Collection m_oProcessingSiteItem_Collection = new ProcessingSiteItem_Collection();
        public TransportationCosts m_oTranCosts = new TransportationCosts();
        public ProcessorScenarioItem_Collection m_oProcessorScenarioItem_Collection = new ProcessorScenarioItem_Collection();
        public ConditionTableSQLFilter m_oCondTableSQLFilter = new ConditionTableSQLFilter();

        public CoreAnalysisScenarioItem()
        {
            
        }
        private string _strScenarioId = "";
        public string ScenarioId
        {
            get { return _strScenarioId; }
            set { _strScenarioId = value; }
        }
        private string _strDesc = "";
        public string Description
        {
            get { return _strDesc; }
            set { _strDesc = value; }
        }
        private string _strDbPath = "";
        public string DbPath
        {
            get { return _strDbPath; }
            set { _strDbPath = value; }
        }
        private string _strDbFileName = "";
        public string DbFileName
        {
            get { return _strDbFileName; }
            set { _strDbFileName = value; }
        }
        private string _strOwnerGroupCdList = "";
        public string OwnerGroupCodeList
        {
            get { return _strOwnerGroupCdList; }
            set { _strOwnerGroupCdList = value; }
        }
        
        
        
        private string _strPlotTableSQLFilter = "";
        public string PlotTableSQLFilter
        {
            get {return _strPlotTableSQLFilter;}
            set {_strPlotTableSQLFilter=value;}
        }
        public struct ConditionTableSQLFilter
        {
            public string SQL;
            public string LowSlopeMaximumYardingDistanceFeet;
            public string SteepSlopeMaximumYardingDistanceFeet;
        }
        public void Copy(CoreAnalysisScenarioItem p_oSource,
                         CoreAnalysisScenarioItem p_oDest)
        {
            p_oDest.DbFileName = p_oSource.DbFileName;
            p_oDest.DbPath = p_oSource.DbPath;
            p_oDest.OwnerGroupCodeList = p_oSource.OwnerGroupCodeList;
            p_oDest.PlotTableSQLFilter = p_oSource.PlotTableSQLFilter;
            p_oDest.ScenarioId = p_oSource.ScenarioId;
            p_oDest.Description = p_oSource.Description;
            p_oDest.m_oEffectiveVariablesItem_Collection.Copy(
                p_oSource.m_oEffectiveVariablesItem_Collection,
            ref p_oDest.m_oEffectiveVariablesItem_Collection, true);
            p_oDest.m_oOptimizationVariableItem_Collection.Copy(
                p_oSource.m_oOptimizationVariableItem_Collection,
            ref p_oDest.m_oOptimizationVariableItem_Collection,true);
            p_oDest.m_oProcessingSiteItem_Collection.Copy(
                p_oSource.m_oProcessingSiteItem_Collection,
            ref p_oDest.m_oProcessingSiteItem_Collection, true);
            p_oDest.m_oRxIntensityItem_Collection.Copy(
                p_oSource.m_oRxIntensityItem_Collection,
            ref p_oDest.m_oRxIntensityItem_Collection, true);
            p_oDest.m_oTieBreaker_Collection.Copy(
                p_oSource.m_oTieBreaker_Collection,
            ref p_oDest.m_oTieBreaker_Collection, true);
            p_oDest.m_oTranCosts.Copy(
                p_oSource.m_oTranCosts,
                p_oDest.m_oTranCosts);
            p_oDest.m_oCondTableSQLFilter = p_oSource.m_oCondTableSQLFilter;
            p_oDest.m_oProcessorScenarioItem_Collection.Copy(
                p_oSource.m_oProcessorScenarioItem_Collection,
            ref p_oDest.m_oProcessorScenarioItem_Collection,true);

            
        }



        public class TransportationCosts
        {

            public string RoadHaulCostPerGreenTonPerHour = "";
            public string RailHaulCostPerGreenTonPerMile = "";
            public string RailChipTransferPerGreenTonPerHour = "";
            public string RailMerchTransferPerGreenTonPerHour = "";

            public void Copy(TransportationCosts p_oSource, TransportationCosts p_oDest)
            {
                p_oDest.RailChipTransferPerGreenTonPerHour = p_oSource.RailChipTransferPerGreenTonPerHour;
                p_oDest.RailHaulCostPerGreenTonPerMile = p_oSource.RailHaulCostPerGreenTonPerMile;
                p_oDest.RailMerchTransferPerGreenTonPerHour = p_oSource.RailMerchTransferPerGreenTonPerHour;
                p_oDest.RoadHaulCostPerGreenTonPerHour = p_oSource.RoadHaulCostPerGreenTonPerHour;
            }
            
        }
        
        public class EffectiveVariablesItem
        {
            public static byte NUMBER_OF_VARIABLES = 4;
            public string[] m_strPreVarArray = new string[NUMBER_OF_VARIABLES];
            public string[] m_strPostVarArray = new string[NUMBER_OF_VARIABLES];
            public string[] m_strBetterExpr = new string[NUMBER_OF_VARIABLES];
            public string[] m_strWorseExpr = new string[NUMBER_OF_VARIABLES];
            public string[] m_strEffectiveExpr = new string[NUMBER_OF_VARIABLES];
            public string m_strOverallEffectiveExpr = "";
            public bool m_bOverallEffectiveNetRevEnabled = false;
            public string m_strOverallEffectiveNetRevOperator = ">";
            public string m_strOverallEffectiveNetRevValue = "0";
            private string _strRxCycle = "";
            private string _strType="";

            public EffectiveVariablesItem()
            {
                for (int x = 0; x <= NUMBER_OF_VARIABLES - 1; x++)
                {
                    m_strPreVarArray[x] = "Not Defined";
                    m_strPostVarArray[x] = "Not Defined";
                    m_strBetterExpr[x] = "";
                    m_strWorseExpr[x] = "";
                    m_strEffectiveExpr[x] = "";
                    _strRxCycle = "";
                    _strType="";
                }
            }
            public string RxCycle
            {
                get { return _strRxCycle; }
                set { _strRxCycle = value; }
            }
            public string Type
            {
                get {return _strType;}
                set {_strType=value;}
            }
            public void Copy(EffectiveVariablesItem p_oSource,
                             EffectiveVariablesItem p_oDest)
            {
                for (int x = 0; x <= NUMBER_OF_VARIABLES - 1; x++)
                {
                    p_oDest.m_strPreVarArray[x] = p_oSource.m_strPreVarArray[x];
                    p_oDest.m_strPostVarArray[x] = p_oSource.m_strPostVarArray[x];
                    p_oDest.m_strBetterExpr[x] = p_oSource.m_strBetterExpr[x];
                    p_oDest.m_strWorseExpr[x] = p_oSource.m_strWorseExpr[x];
                    p_oDest.m_strEffectiveExpr[x] = p_oSource.m_strEffectiveExpr[x];

                }
                p_oDest.RxCycle = p_oSource.RxCycle;
                p_oDest.m_strOverallEffectiveExpr = p_oSource.m_strOverallEffectiveExpr;
                p_oDest.m_bOverallEffectiveNetRevEnabled = p_oSource.m_bOverallEffectiveNetRevEnabled;
                p_oDest.m_strOverallEffectiveNetRevOperator = p_oSource.m_strOverallEffectiveNetRevOperator;
                p_oDest.m_strOverallEffectiveNetRevValue = p_oSource.m_strOverallEffectiveNetRevValue;
                
            }
            /// <summary>
            /// return the table names found in the either m_strPreVarArray or m_strPostVarArray variables
            /// </summary>
            /// <param name="p_oVarArray">m_strPreVarArray or m_strPostVarArray values</param>
            /// <returns></returns>
            public string[] TableNames(string[] p_oVarArray)
            {
                string strTable = "";
                string strTableList = ",";
                string[] strTableArray = null;
                for (int x = 0; x <= p_oVarArray.Length - 1; x++)
                {
                    strTable = this.TableName(p_oVarArray[x]);
                    if (strTable.Trim().Length > 0)
                    {
                        if (strTableList.Trim().ToUpper().IndexOf("," + strTable.ToUpper().Trim().ToUpper() + ",", 0) != 0)
                        {
                            strTableList = strTableList + strTable.Trim() + ",";
                        }

                    }
                }

                if (strTableList.Trim().Length > 1)
                {
                    strTableList = strTableList.Substring(1, strTableList.Length - 2);
                    strTableArray = frmMain.g_oUtils.ConvertListToArray(strTableList, ",");

                }
                else
                {
                    strTableList = "";
                }
                return strTableArray;
            }
            /// <summary>
            /// expecting a value in the format of tablename.columnname
            /// </summary>
            /// <param name="p_strValue"></param>
            /// <returns></returns>
            public string TableName(string p_strValue)
            {
                string strTableName = "";
                if (p_strValue.Trim().ToUpper() != "NOT DEFINED")
                {
                    if (p_strValue.IndexOf(".", 0) > 0)
                    {
                        string[] strArray = frmMain.g_oUtils.ConvertListToArray(p_strValue, ".");
                        if (strArray[0].Trim().Length > 0)
                        {
                            strTableName = strArray[0].Trim();
                        }
                    }
                }
                return strTableName;
            }
            /// <summary>
            /// expecting a value in the format of tablename.columnname
            /// </summary>
            /// <param name="p_strValue"></param>
            /// <returns></returns>
            public string ColumnName(string p_strValue)
            {
                string strColumnName = "";
                if (p_strValue.Trim().ToUpper() != "NOT DEFINED")
                {
                    if (p_strValue.IndexOf(".", 0) > 0)
                    {
                        string[] strArray = frmMain.g_oUtils.ConvertListToArray(p_strValue, ".");
                        if (strArray.Length == 2)
                        {
                            if (strArray[1].Trim().Length > 0)
                            {
                                strColumnName = strArray[1].Trim();
                            }
                        }
                    }
                }
                return strColumnName;
            }


            public bool Modified(EffectiveVariablesItem p_oCompare)
            {
                if (m_strOverallEffectiveExpr.Trim().ToUpper() != p_oCompare.m_strOverallEffectiveExpr.Trim().ToUpper()) return true;
                if (m_strOverallEffectiveNetRevValue.Trim() != p_oCompare.m_strOverallEffectiveNetRevValue.Trim()) return true;
                if (m_bOverallEffectiveNetRevEnabled != p_oCompare.m_bOverallEffectiveNetRevEnabled) return true;
                if (m_strOverallEffectiveNetRevOperator.Trim() != p_oCompare.m_strOverallEffectiveNetRevOperator) return true;


                for (int x = 0; x <= NUMBER_OF_VARIABLES - 1; x++)
                {
                    if (m_strPreVarArray[x].Trim().ToUpper() != p_oCompare.m_strPreVarArray[x].Trim().ToUpper())
                        return true;

                    if (m_strPostVarArray[x].Trim().ToUpper() != p_oCompare.m_strPostVarArray[x].Trim().ToUpper())
                        return true;

                    if (m_strBetterExpr[x].Trim().ToUpper() != p_oCompare.m_strBetterExpr[x].Trim().ToUpper())
                        return true;

                    if (this.m_strWorseExpr[x].Trim().ToUpper() != p_oCompare.m_strWorseExpr[x].Trim().ToUpper())
                        return true;

                    if (this.m_strEffectiveExpr[x].Trim().ToUpper() != p_oCompare.m_strEffectiveExpr[x].Trim().ToUpper())
                        return true;



                }
                return false;
            }
        }
        public class EffectiveVariablesItem_Collection : System.Collections.CollectionBase
        {
            public EffectiveVariablesItem_Collection()
            {
                //
                // TODO: Add constructor logic here
                //
            }

            public void Add(EffectiveVariablesItem m_EffectiveVariablesItem)
            {
                // vérify if object is not already in
                if (this.List.Contains(m_EffectiveVariablesItem))
                    throw new InvalidOperationException();

                // adding it
                this.List.Add(m_EffectiveVariablesItem);

                // return collection
                //return this;
            }
            public void Remove(int index)
            {
                // Check to see if there is a widget at the supplied index.
                if (index > Count - 1 || index < 0)
                // If no widget exists, a messagebox is shown and the operation 
                // is canColumned.
                {
                    System.Windows.Forms.MessageBox.Show("Index not valid!");
                }
                else
                {
                    List.RemoveAt(index);
                }
            }
            public EffectiveVariablesItem Item(int Index)
            {
                // The appropriate item is retrieved from the List object and
                // explicitly cast to the Widget type, then returned to the 
                // caller.
                return (EffectiveVariablesItem)List[Index];
            }
            public void Copy(EffectiveVariablesItem_Collection p_oSource,
                         ref EffectiveVariablesItem_Collection p_oDest, bool p_bInitializeDest)
            {
                int x;
                if (p_bInitializeDest) p_oDest.Clear();
                for (x = 0; x <= p_oSource.Count - 1; x++)
                {
                    EffectiveVariablesItem oItem = new EffectiveVariablesItem();
                    oItem.Copy(p_oSource.Item(x), oItem);
                    p_oDest.Add(oItem);

                }
            }

        }
        public class OptimizationVariableItem
		{
			public bool bSelected=false;
			public string strOptimizedVariable="";
			public string strFVSVariableName="";
			public string strValueSource="";
			public string strMaxYN="N";
			public string strMinYN="N";
			public bool bUseFilter=true;
			public string strFilterOperator="";
			public double dblFilterValue=0;
			public int intListViewIndex=-1;
            private string _strRxCycle = "";
            private string _strType = "";

            public string RxCycle
            {
                get { return _strRxCycle; }
                set { _strRxCycle = value; }
            }
            public string Type
            {
                get { return _strType; }
                set { _strType = value; }
            }
			public void Copy(OptimizationVariableItem p_oSource,
                         ref OptimizationVariableItem p_oDest)
			{
				p_oDest.bSelected = p_oSource.bSelected;
				p_oDest.strOptimizedVariable=p_oSource.strOptimizedVariable;
				p_oDest.strValueSource = p_oSource.strValueSource;
				p_oDest.strMaxYN = p_oSource.strMaxYN;
				p_oDest.strMinYN = p_oSource.strMinYN;
				p_oDest.strFilterOperator = p_oSource.strFilterOperator;
				p_oDest.strFVSVariableName = p_oSource.strFVSVariableName;
				p_oDest.dblFilterValue = p_oSource.dblFilterValue;
				p_oDest.bUseFilter = p_oSource.bUseFilter;
				p_oDest.intListViewIndex = p_oSource.intListViewIndex;
                p_oDest.RxCycle = p_oSource.RxCycle;
                p_oDest.Type = p_oSource.Type;
			}
		}

		public class OptimizationVariableItem_Collection : System.Collections.CollectionBase
		{
			public OptimizationVariableItem_Collection()
			{
				//
				// TODO: Add constructor logic here
				//
			}

			public void Add(OptimizationVariableItem m_oVariable)
			{
				// vérify if object is not already in
				if (this.List.Contains(m_oVariable))
					throw new InvalidOperationException();
 
				// adding it
				this.List.Add(m_oVariable);
 
				// return collection
				//return this;
			}
			public void Remove(int index)
			{
				// Check to see if there is a widget at the supplied index.
				if (index > Count - 1 || index < 0)
					// If no widget exists, a messagebox is shown and the operation 
					// is canColumned.
				{
					System.Windows.Forms.MessageBox.Show("Index not valid!");
				}
				else
				{
					List.RemoveAt(index); 
				}
			}
			public OptimizationVariableItem Item(int Index)
			{
				// The appropriate item is retrieved from the List object and
				// explicitly cast to the Widget type, then returned to the 
				// caller.
				return (OptimizationVariableItem) List[Index];
			}
			public void Copy(OptimizationVariableItem_Collection p_oSource,
			             ref OptimizationVariableItem_Collection p_oDest,bool p_bInitializeDest)
			{
				int x;
				if (p_bInitializeDest) p_oDest.Clear();
				for (x=0;x<=p_oSource.Count-1;x++)
				{
					OptimizationVariableItem oItem = new OptimizationVariableItem();
					oItem.Copy(p_oSource.Item(x),ref oItem);
					p_oDest.Add(oItem);

				}
			}


		}
        public class TieBreakerItem
        {
            public bool bSelected = true;
            public string strMethod = "";
            public string strFVSVariableName = "";
            public string strMaxYN = "N";
            public string strMinYN = "N";
            public string strValueSource = "";
            public int intListViewIndex = -1;
            private string _strRxCycle = "";

            public string RxCycle
            {
                get { return _strRxCycle; }
                set { _strRxCycle = value; }
            }
            public void Copy(TieBreakerItem p_oSource, ref TieBreakerItem p_oDest)
            {
                p_oDest.bSelected = p_oSource.bSelected;
                p_oDest.strMethod = p_oSource.strMethod;
                p_oDest.strFVSVariableName = p_oSource.strFVSVariableName;
                p_oDest.strMaxYN = p_oSource.strMaxYN;
                p_oDest.strMinYN = p_oSource.strMinYN;
                p_oDest.strValueSource = p_oSource.strValueSource;
                p_oDest.intListViewIndex = p_oSource.intListViewIndex;
                p_oDest.RxCycle = p_oSource.RxCycle;


            }
        }

        public class TieBreaker_Collection : System.Collections.CollectionBase
        {
            public TieBreaker_Collection()
            {
                //
                // TODO: Add constructor logic here
                //
            }

            public void Add(TieBreakerItem m_oTieBreaker)
            {
                // vérify if object is not already in
                if (this.List.Contains(m_oTieBreaker))
                    throw new InvalidOperationException();

                // adding it
                this.List.Add(m_oTieBreaker);

                // return collection
                //return this;
            }
            public void Remove(int index)
            {
                // Check to see if there is a widget at the supplied index.
                if (index > Count - 1 || index < 0)
                // If no widget exists, a messagebox is shown and the operation 
                // is canColumned.
                {
                    System.Windows.Forms.MessageBox.Show("Index not valid!");
                }
                else
                {
                    List.RemoveAt(index);
                }
            }
            public TieBreakerItem Item(int Index)
            {
                // The appropriate item is retrieved from the List object and
                // explicitly cast to the Widget type, then returned to the 
                // caller.
                return (TieBreakerItem)List[Index];
            }
            public void Copy(TieBreaker_Collection p_oSource,
                         ref TieBreaker_Collection p_oDest, bool p_bInitializeDest)
            {
                int x;
                if (p_bInitializeDest) p_oDest.Clear();
                for (x = 0; x <= p_oSource.Count - 1; x++)
                {
                    TieBreakerItem oItem = new TieBreakerItem();
                    oItem.Copy(p_oSource.Item(x), ref oItem);
                    p_oDest.Add(oItem);

                }
            }


        }
        public class RxIntensityItem
        {

            



           
            private string _strRx = "";
            public string Rx
            {
                get { return _strRx; }
                set { _strRx = value; }
            }
            private int _intRxIntensity = -1;
            public int RxIntensity
            {
                get { return _intRxIntensity; }
                set { _intRxIntensity = value; }
            }
            public void Copy(RxIntensityItem p_oSource, ref RxIntensityItem p_oDest)
            {
                p_oDest.Rx = p_oSource.Rx;
                p_oDest.RxIntensity = p_oSource.RxIntensity;
                

            }
        }

        public class RxIntensityItem_Collection : System.Collections.CollectionBase
        {
            public RxIntensityItem_Collection()
            {
                //
                // TODO: Add constructor logic here
                //
            }

            public void Add(RxIntensityItem m_oRxIntensityItem)
            {
                // vérify if object is not already in
                if (this.List.Contains(m_oRxIntensityItem))
                    throw new InvalidOperationException();

                // adding it
                this.List.Add(m_oRxIntensityItem);

                // return collection
                //return this;
            }
            public void Remove(int index)
            {
                // Check to see if there is a widget at the supplied index.
                if (index > Count - 1 || index < 0)
                // If no widget exists, a messagebox is shown and the operation 
                // is canColumned.
                {
                    System.Windows.Forms.MessageBox.Show("Index not valid!");
                }
                else
                {
                    List.RemoveAt(index);
                }
            }
            public RxIntensityItem Item(int Index)
            {
                // The appropriate item is retrieved from the List object and
                // explicitly cast to the Widget type, then returned to the 
                // caller.
                return (RxIntensityItem)List[Index];
            }
            public void Copy(RxIntensityItem_Collection p_oSource,
                         ref RxIntensityItem_Collection p_oDest, bool p_bInitializeDest)
            {
                int x;
                if (p_bInitializeDest) p_oDest.Clear();
                for (x = 0; x <= p_oSource.Count - 1; x++)
                {
                    RxIntensityItem oItem = new RxIntensityItem();
                    oItem.Copy(p_oSource.Item(x), ref oItem);
                    p_oDest.Add(oItem);

                }
            }


        }
        public class ProcessingSiteItem
        {
            public int intListViewIndex = -1;
            string _strPSiteId = "";
            public string ProcessingSiteId
            {
                get { return _strPSiteId; }
                set { _strPSiteId = value; }
            }
            string _strPSiteName = "";
            public string ProcessingSiteName
            {
                get { return _strPSiteName; }
                set { _strPSiteName = value; }
            }
            public string[,] m_strTranCdDescArray = new string[3, 2] {{"1","Regular - PSite With Road Only Access"},
		  															  {"2","Railhead - Road To Rail Wood Transfer Point"},
		  															  {"3","Rail Collector - PSite With Both Road And Rail Access"}
		  															 };
            string _strTranCd = "";
            public string TransportationCode
            {
                get { return _strTranCd; }
                set { _strTranCd = value; }
            }
            public string[,] m_strBioCdDescArray = new string[3, 2] {{"1","Merchantable - Logs Only"},
						  											 {"2","Chips - Chips Only"},
							  										 {"3","Both - Logs And Chips"}
							  										};
            string _strBioCd = "";
            public string BiomassCode
            {
                get { return _strBioCd; }
                set { _strBioCd = value; }
            }

            string _strExistsYN = "N";
            public string ProcessingSiteExistYN
            {
                get { return _strExistsYN; }
                set { _strExistsYN = value; }
            }

            double _dblLat = -1;
            public double GPSLatitude
            {
                get { return _dblLat; }
                set { _dblLat = value; }
            }

            double _dblLon = -1;
            public double GPSLongitude
            {
                get { return _dblLon; }
                set { _dblLon = value; }
            }
            bool _bSelected = false;
            public bool Selected
            {
                get { return _bSelected; }
                set { _bSelected = value; }
            }




            public void Copy(ProcessingSiteItem p_oSource, ref ProcessingSiteItem p_oDest)
            {
                p_oDest.ProcessingSiteId = p_oSource.ProcessingSiteId;
                p_oDest.ProcessingSiteName = p_oSource.ProcessingSiteName;
                p_oDest.TransportationCode = p_oSource.TransportationCode;
                p_oDest.BiomassCode = p_oSource.BiomassCode;
                p_oDest.ProcessingSiteExistYN = p_oSource.ProcessingSiteExistYN;
                p_oDest.GPSLatitude = p_oSource.GPSLatitude;
                p_oDest.GPSLongitude = p_oSource.GPSLongitude;
                p_oDest.Selected = p_oSource.Selected;
                p_oDest.intListViewIndex = p_oSource.intListViewIndex;
            }
        }

        public class ProcessingSiteItem_Collection : System.Collections.CollectionBase
        {
            public ProcessingSiteItem_Collection()
            {
                //
                // TODO: Add constructor logic here
                //
            }

            public void Add(ProcessingSiteItem m_oPSite)
            {
                // vérify if object is not already in
                if (this.List.Contains(m_oPSite))
                    throw new InvalidOperationException();

                // adding it
                this.List.Add(m_oPSite);

                // return collection
                //return this;
            }
            public void Remove(int index)
            {
                // Check to see if there is a widget at the supplied index.
                if (index > Count - 1 || index < 0)
                // If no widget exists, a messagebox is shown and the operation 
                // is canColumned.
                {
                    System.Windows.Forms.MessageBox.Show("Index not valid!");
                }
                else
                {
                    List.RemoveAt(index);
                }
            }
            public ProcessingSiteItem Item(int Index)
            {
                // The appropriate item is retrieved from the List object and
                // explicitly cast to the Widget type, then returned to the 
                // caller.
                return (ProcessingSiteItem)List[Index];
            }
            public void Copy(ProcessingSiteItem_Collection p_oSource,
                         ref ProcessingSiteItem_Collection p_oDest, bool p_bInitializeDest)
            {
                int x;
                if (p_bInitializeDest) p_oDest.Clear();
                for (x = 0; x <= p_oSource.Count - 1; x++)
                {
                    ProcessingSiteItem oItem = new ProcessingSiteItem();
                    oItem.Copy(p_oSource.Item(x), ref oItem);
                    p_oDest.Add(oItem);

                }
            }


        }
    }
    public class CoreAnalysisScenarioItem_Collection : System.Collections.CollectionBase
    {
        public CoreAnalysisScenarioItem_Collection()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        public void Add(FIA_Biosum_Manager.CoreAnalysisScenarioItem m_CoreAnalysisScenarioItem)
        {
            // vérify if object is not already in
            if (this.List.Contains(m_CoreAnalysisScenarioItem))
                throw new InvalidOperationException();

            // adding it
            this.List.Add(m_CoreAnalysisScenarioItem);

            // return collection
            //return this;
        }
        public void Remove(int index)
        {
            // Check to see if there is a widget at the supplied index.
            if (index > Count - 1 || index < 0)
            // If no widget exists, a messagebox is shown and the operation 
            // is canColumned.
            {
                System.Windows.Forms.MessageBox.Show("Index not valid!");
            }
            else
            {
                List.RemoveAt(index);
            }
        }
        public FIA_Biosum_Manager.CoreAnalysisScenarioItem Item(int Index)
        {
            // The appropriate item is retrieved from the List object and
            // explicitly cast to the Widget type, then returned to the 
            // caller.
            return (FIA_Biosum_Manager.CoreAnalysisScenarioItem)List[Index];
        }

    }




    public class CoreAnalysisScenarioTools
    {
        public string m_strError = "";
        public int m_intError = 0;
        public CoreAnalysisScenarioTools()
        {
        }
        public void LoadScenario(string p_strScenarioId, Queries p_oQueries,CoreAnalysisScenarioItem_Collection p_oCoreAnalysisScenarioItem_Collection)
        {
           
            //
            //LOAD PROJECT DATATASOURCES INFO
            //
            p_oQueries.m_oFvs.LoadDatasource = true;
            p_oQueries.m_oFIAPlot.LoadDatasource = true;
            p_oQueries.m_oProcessor.LoadDatasource = true;
            p_oQueries.m_oReference.LoadDatasource = true;
            p_oQueries.LoadDatasources(true, "core", p_strScenarioId);
            p_oQueries.m_oDataSource.CreateScenarioRuleDefinitionTableLinks(
                p_oQueries.m_strTempDbFile,
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim(),
                "C");
            LoadAll(p_oQueries.m_strTempDbFile, p_oQueries, p_strScenarioId, p_oCoreAnalysisScenarioItem_Collection);
        }
        public void LoadAll(string p_strDbFile, Queries p_oQueries, string p_strScenarioId, FIA_Biosum_Manager.CoreAnalysisScenarioItem_Collection p_oCoreAnalysisScenarioItem_Collection)
        {
            int x;
            for (x = p_oCoreAnalysisScenarioItem_Collection.Count - 1; x >= 0; x--)
            {
                if (p_oCoreAnalysisScenarioItem_Collection.Item(x).ScenarioId.Trim().ToUpper() ==
                    p_strScenarioId.Trim().ToUpper())
                {
                    p_oCoreAnalysisScenarioItem_Collection.Remove(x);
                }
            }
            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(p_strDbFile, "", ""));
            if (oAdo.m_intError == 0)
            {

                CoreAnalysisScenarioItem oItem = new CoreAnalysisScenarioItem();
                this.LoadGeneral(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                this.LoadEffectiveVariables(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                this.LoadOptimizationVariable(oAdo, oAdo.m_OleDbConnection, p_strScenarioId,oItem);
                this.LoadTieBreakerVariables(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                this.LoadRxIntensity(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                this.LoadProcessorScenarioItems(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                this.LoadPlotFilter(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                this.LoadCondFilter(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                this.LoadProcessingSites(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                this.LoadLandOwnerGroupFilter(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                this.LoadTransportationCosts(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                p_oCoreAnalysisScenarioItem_Collection.Add(oItem);
                




            }
            m_intError = oAdo.m_intError;
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oAdo = null;
        }

        public void LoadGeneral(string p_strDbFile, string p_strScenarioId, FIA_Biosum_Manager.CoreAnalysisScenarioItem p_oCoreAnalysisScenarioItem)
        {
            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(p_strDbFile, "", ""));
            if (oAdo.m_intError == 0)
            {
                this.LoadGeneral(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, p_oCoreAnalysisScenarioItem);
            }
            m_intError = oAdo.m_intError;
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oAdo = null;
        }
        public void LoadGeneral(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.CoreAnalysisScenarioItem p_oCoreAnalysisScenarioItem)
        {
            p_oAdo.SqlQueryReader(p_oConn, "SELECT * FROM scenario " + " WHERE TRIM(UCASE(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "'");

            if (p_oAdo.m_intError == 0)
            {
                if (p_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (p_oAdo.m_OleDbDataReader.Read())
                    {
                        //
                        //SCENARIO ID
                        //
                        if (p_oAdo.m_OleDbDataReader["scenario_id"] != System.DBNull.Value)
                        {
                            p_oCoreAnalysisScenarioItem.ScenarioId = p_oAdo.m_OleDbDataReader["scenario_id"].ToString().Trim();
                        }
                        //
                        //DESCRIPTION
                        //
                        if (p_oAdo.m_OleDbDataReader["description"] != System.DBNull.Value)
                        {
                            p_oCoreAnalysisScenarioItem.Description = p_oAdo.m_OleDbDataReader["description"].ToString().Trim();
                        }
                        //
                        //PATH
                        //
                        if (p_oAdo.m_OleDbDataReader["path"] != System.DBNull.Value)
                        {
                            p_oCoreAnalysisScenarioItem.DbPath = p_oAdo.m_OleDbDataReader["path"].ToString().Trim();
                        }
                        //
                        //FILE
                        //
                        if (p_oAdo.m_OleDbDataReader["file"] != System.DBNull.Value)
                        {
                            p_oCoreAnalysisScenarioItem.DbFileName = p_oAdo.m_OleDbDataReader["file"].ToString().Trim();
                        }


                    }
                    p_oAdo.m_OleDbDataReader.Close();
                }
            }
        }
        public void LoadEffectiveVariables(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.CoreAnalysisScenarioItem p_oCoreAnalysisScenarioItem)
        {
            int x = 0, y = 0;
            string strRxCycle = "";
            //
            //INITIALIZE EFFECTIVE VARIABLES COLLECTION
            //
            for (x = p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Count - 1;x>=0 ; x--)
            {
                p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Remove(x);
            }
            //
            //LOAD EFFECTIVE VARIABLES
            //
            p_oAdo.m_strSQL = "SELECT a.* " +
                            "FROM scenario_fvs_variables a " +
                            "WHERE TRIM(a.scenario_id)='" + p_strScenarioId.Trim() + "' AND " +
                            "a.current_yn='Y' ORDER BY a.rxcycle,a.variable_number";

            p_oAdo.SqlQueryReader(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows)
            {
                    
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                    int intVarNum = Convert.ToInt32(p_oAdo.m_OleDbDataReader["variable_number"]) - 1;
                    strRxCycle = Convert.ToString(p_oAdo.m_OleDbDataReader["rxcycle"]).Trim();
                    //find the current rxcycle in the collection
                    for (x = 0; x <= p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Count - 1; x++)
                    {
                        if (p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).RxCycle == strRxCycle)
                            break;
                    }

                    if (x > p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Count - 1)
                    {
                        CoreAnalysisScenarioItem.EffectiveVariablesItem oEffItem = new CoreAnalysisScenarioItem.EffectiveVariablesItem();
                        oEffItem.RxCycle = strRxCycle;
                        p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Add(oEffItem);
                    }
                    p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strPreVarArray[intVarNum] =
                        Convert.ToString(p_oAdo.m_OleDbDataReader["pre_fvs_variable"]).Trim();

                    p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strPostVarArray[intVarNum] =
                        Convert.ToString(p_oAdo.m_OleDbDataReader["post_fvs_variable"]).Trim();


                    if (p_oAdo.m_OleDbDataReader["better_expression"] != System.DBNull.Value)
                        p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strBetterExpr[intVarNum] =
                            Convert.ToString(p_oAdo.m_OleDbDataReader["better_expression"]).Trim();

                    if (p_oAdo.m_OleDbDataReader["worse_expression"] != System.DBNull.Value)
                        p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strWorseExpr[intVarNum] =
                                Convert.ToString(p_oAdo.m_OleDbDataReader["worse_expression"]).Trim();

                    if (p_oAdo.m_OleDbDataReader["effective_expression"] != System.DBNull.Value)
                        p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strEffectiveExpr[intVarNum] =
                            Convert.ToString(p_oAdo.m_OleDbDataReader["effective_expression"]).Trim();

                    if (p_oAdo.m_OleDbDataReader["rxcycle"] != System.DBNull.Value)
                        p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).RxCycle = 
                            Convert.ToString(p_oAdo.m_OleDbDataReader["rxcycle"]).Trim();


                }
            }

            p_oAdo.m_OleDbDataReader.Close();
            //
            //LOAD OVERALL EFFECTIVE VARIABLES
            //
            //overall expression
            p_oAdo.m_strSQL = "SELECT b.overall_effective_expression,b.current_yn," +
                                   "b.nr_dpa_filter_enabled_yn,b.nr_dpa_filter_operator," +
                                   "b.nr_dpa_filter_value,b.rxcycle " +
                            "FROM scenario_fvs_variables_overall_effective b " +
                            "WHERE TRIM(b.scenario_id)='" + p_strScenarioId.Trim() + "' AND " +
                            "b.current_yn='Y'";



            p_oAdo.SqlQueryReader(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows)
            {
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                    strRxCycle = Convert.ToString(p_oAdo.m_OleDbDataReader["rxcycle"]).Trim();
                    //find the current rxcycle in the collection
                    for (x = 0; x <= p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Count - 1; x++)
                    {
                        if (p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).RxCycle == strRxCycle)
                            break;
                    }
                    if (p_oAdo.m_OleDbDataReader["overall_effective_expression"] != System.DBNull.Value &&
                        p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strOverallEffectiveExpr.Trim().Length == 0)
                        p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strOverallEffectiveExpr =
                                Convert.ToString(p_oAdo.m_OleDbDataReader["overall_effective_expression"]).Trim();

                    //enable filter
                    if (p_oAdo.m_OleDbDataReader["nr_dpa_filter_enabled_yn"] != System.DBNull.Value)
                    {
                        if (Convert.ToString(p_oAdo.m_OleDbDataReader["nr_dpa_filter_enabled_yn"]).Trim() == "Y")
                            p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_bOverallEffectiveNetRevEnabled = true;
                        else
                            p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_bOverallEffectiveNetRevEnabled = false;
                    }
                    else
                    {
                        p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_bOverallEffectiveNetRevEnabled = false;
                    }
                    //filter operator
                    if (p_oAdo.m_OleDbDataReader["nr_dpa_filter_operator"] != System.DBNull.Value)
                    {

                        p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strOverallEffectiveNetRevOperator = Convert.ToString(p_oAdo.m_OleDbDataReader["nr_dpa_filter_operator"]).Trim();
                    }
                    else
                    {
                        p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strOverallEffectiveNetRevOperator = ">";
                    }
                    //filter value
                    if (p_oAdo.m_OleDbDataReader["nr_dpa_filter_value"] != System.DBNull.Value)
                    {
                        p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strOverallEffectiveNetRevValue = Convert.ToString(p_oAdo.m_OleDbDataReader["nr_dpa_filter_value"]).Trim();
                    }
                    else p_oCoreAnalysisScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strOverallEffectiveNetRevValue = "0";

                    


                }
            }

            p_oAdo.m_OleDbDataReader.Close();
            
            

               
        }
        public void LoadOptimizationVariable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.CoreAnalysisScenarioItem p_oCoreAnalysisScenarioItem)
        {
            int x = 0, y = 0;
            int intVarNum = 0;

            //
            //INITIALIZE OPTIMIZATION VARIABLES COLLECTION
            //
            for (x = p_oCoreAnalysisScenarioItem.m_oOptimizationVariableItem_Collection.Count - 1;x>=0 ; x--)
            {
                p_oCoreAnalysisScenarioItem.m_oOptimizationVariableItem_Collection.Remove(x);
            }
            
            p_oAdo.m_strSQL = "SELECT * " +
                "FROM scenario_fvs_variables_optimization " +
                "WHERE TRIM(scenario_id)='" + p_strScenarioId.Trim() + "' AND " +
                "current_yn='Y'";

            p_oAdo.SqlQueryReader(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows)
            {
                
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                   

                    if (p_oAdo.m_OleDbDataReader["optimization_variable"] != System.DBNull.Value)
                    {
                        CoreAnalysisScenarioItem.OptimizationVariableItem oItem = new CoreAnalysisScenarioItem.OptimizationVariableItem();
                        oItem.strOptimizedVariable = Convert.ToString(p_oAdo.m_OleDbDataReader["optimization_variable"]);
                        //optimization variable
                        if (p_oAdo.m_OleDbDataReader["fvs_variable_name"] != System.DBNull.Value)
                        {
                            oItem.strFVSVariableName = Convert.ToString(p_oAdo.m_OleDbDataReader["fvs_variable_name"]).Trim();
                        }
                        else
                        {
                            oItem.strFVSVariableName = "NA";
                        }
                        //value source (POST or POST-PRE)
                        if (p_oAdo.m_OleDbDataReader["value_source"] != System.DBNull.Value)
                        {
                            oItem.strValueSource = Convert.ToString(p_oAdo.m_OleDbDataReader["value_source"]).Trim();
                        }
                        else
                        {
                            oItem.strValueSource = "Not Defined";
                        }

                        //max value
                        if (p_oAdo.m_OleDbDataReader["max_yn"] != System.DBNull.Value)
                        {
                            oItem.strMaxYN = Convert.ToString(p_oAdo.m_OleDbDataReader["max_yn"]).Trim();
                        }
                        else
                        {
                            oItem.strMaxYN = "N";
                        }
                        //min value
                        if (p_oAdo.m_OleDbDataReader["min_yn"] != System.DBNull.Value)
                        {
                            oItem.strMinYN = Convert.ToString(p_oAdo.m_OleDbDataReader["min_yn"]).Trim();
                        }
                        else
                        {
                            oItem.strMinYN = "N";
                        }
                        //enable filter
                        if (p_oAdo.m_OleDbDataReader["filter_enabled_yn"] != System.DBNull.Value)
                        {
                            if (Convert.ToString(p_oAdo.m_OleDbDataReader["filter_enabled_yn"]).Trim() == "Y")
                                oItem.bUseFilter = true;
                            else
                                oItem.bUseFilter = false;

                        }
                        else
                        {
                            oItem.bUseFilter = false;
                        }
                        //filter operator
                        if (p_oAdo.m_OleDbDataReader["filter_operator"] != System.DBNull.Value)
                        {

                            oItem.strFilterOperator = Convert.ToString(p_oAdo.m_OleDbDataReader["filter_operator"]).Trim();
                        }
                        else
                        {
                            oItem.strFilterOperator = "";
                        }
                        //filter value
                        if (p_oAdo.m_OleDbDataReader["filter_value"] != System.DBNull.Value)
                        {
                            oItem.dblFilterValue = Convert.ToDouble(p_oAdo.m_OleDbDataReader["filter_value"]);
                        }
                        //filter operator
                        if (p_oAdo.m_OleDbDataReader["checked_yn"] != System.DBNull.Value)
                        {
                            if (Convert.ToString(p_oAdo.m_OleDbDataReader["checked_yn"]).Trim() == "Y")
                            {
                                oItem.bSelected = true;
                                if (oItem.strOptimizedVariable.Trim().ToUpper() == "FVS VARIABLE")
                                {
                                }
                                else
                                {
                                }
                            }
                            else
                                oItem.bSelected = false;
                        }
                        else
                        {
                            oItem.bSelected = false;
                        }
                        oItem.RxCycle = p_oAdo.m_OleDbDataReader["rxcycle"].ToString().Trim();
                        p_oCoreAnalysisScenarioItem.m_oOptimizationVariableItem_Collection.Add(oItem);
                    }
                }
                
            }
            p_oAdo.m_OleDbDataReader.Close();

        }
        public void LoadTieBreakerVariables(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.CoreAnalysisScenarioItem p_oCoreAnalysisScenarioItem)
        {
            this.m_intError = 0;
            this.m_strError = "";
            

            int x, y;
            //
            //INITIALIZE TIEBREAKER COLLECTION
            //
            p_oCoreAnalysisScenarioItem.m_oTieBreaker_Collection.Clear();
            //
            //LOAD TIEBREAKER ITEMS
            //
            p_oAdo.m_strSQL = "SELECT * FROM " +
                                Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " " +
                                "WHERE TRIM(scenario_id)='" + p_strScenarioId.Trim() + "'";


            p_oAdo.SqlQueryReader(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows)
            {
                while (p_oAdo.m_OleDbDataReader.Read())
                {


                    if (p_oAdo.m_OleDbDataReader["tiebreaker_method"] != System.DBNull.Value)
                    {
                        CoreAnalysisScenarioItem.TieBreakerItem oItem = new CoreAnalysisScenarioItem.TieBreakerItem();
                        oItem.RxCycle = p_oAdo.m_OleDbDataReader["rxcycle"].ToString().Trim();
                        oItem.strMethod = p_oAdo.m_OleDbDataReader["tiebreaker_method"].ToString().Trim();
                        if (oItem.strMethod.ToUpper() == "FVS VARIABLE")
                        {


                            //fvs variable name
                            oItem.strFVSVariableName = p_oAdo.m_OleDbDataReader["fvs_variable_name"].ToString().Trim();
                            oItem.strValueSource = p_oAdo.m_OleDbDataReader["value_source"].ToString().Trim();



                            //MAX or MIN	
                            if (p_oAdo.m_OleDbDataReader["max_yn"].ToString().Trim().ToUpper() == "Y")
                            {
                                oItem.strMaxYN = "Y";
                                oItem.strMinYN = "N";
                            }
                            else
                            {
                                oItem.strMinYN = "Y";
                                oItem.strMaxYN = "N";

                            }
                            if (p_oAdo.m_OleDbDataReader["checked_yn"].ToString().Trim().ToUpper() == "Y")
                            {
                                oItem.bSelected = true;

                            }
                            else
                            {
                                oItem.bSelected = false;
                            }
                        }
                        else if (oItem.strMethod.ToUpper() == "TREATMENT INTENSITY")
                        {
                            if (p_oAdo.m_OleDbDataReader["checked_yn"].ToString().Trim().ToUpper() == "Y")
                            {
                                oItem.bSelected = true;

                            }
                            else
                            {
                                oItem.bSelected = false;
                            }
                        }
                        p_oCoreAnalysisScenarioItem.m_oTieBreaker_Collection.Add(oItem);
                    }

                }
            }

            p_oAdo.m_OleDbDataReader.Close();
            

        }
        public void LoadRxIntensity(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.CoreAnalysisScenarioItem p_oCoreAnalysisScenarioItem)
        {
            this.m_intError = 0;
            this.m_strError = "";


            int x, y;
            //
            //GET ALL THE CURRENT TREATMENTS
            //
            /*************************************************************************
			 **get the treatment prescription mdb file,table, and connection strings
			 *************************************************************************/
            string strRxMDBFile = "", strRxTableName = "", strRxConn = "";
            p_oAdo.getScenarioDataSourceConnStringAndTable(ref strRxMDBFile,
                                            ref strRxTableName, ref strRxConn,
                                            "Treatment Prescriptions",
                                            p_strScenarioId,
                                            p_oConn);
            //
            //GET A LIST OF ALL THE CURRENT TREATMENTS
            //
            System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
            p_oAdo.OpenConnection(strRxConn,ref oConn);
            string strRxList = p_oAdo.CreateCommaDelimitedList(oConn, "SELECT RX FROM " + strRxTableName, "'");
            string[] strRxArray = frmMain.g_oUtils.ConvertListToArray(strRxList, ",");
            p_oAdo.CloseConnection(oConn);
            //
            //INITIALIZE RXINTENSITY COLLECTION
            //
            p_oCoreAnalysisScenarioItem.m_oRxIntensityItem_Collection.Clear();
            //
            //FIRST DELETE THE RX'S THAT NO LONGER EXIST
            //
            p_oAdo.m_strSQL = "DELETE FROM scenario_rx_intensity WHERE TRIM(UCASE(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "' AND rx NOT IN (" + strRxList + ")";
            p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
            //
            //NEXT INSERT INTO THE SCENARIO_RX_INTENSITY THOSE RX'S THAT DO NOT CURRENTLY EXIST
            //
            if (p_oAdo.TableExist(p_oConn, "rxtemp"))
                p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE rxtemp");
            p_oAdo.SqlNonQuery(p_oConn, Tables.CoreScenarioRuleDefinitions.CreateScenarioRxIntensityTableSQL("rxtemp"));
            for (x = 0; x <= strRxArray.Length - 1; x++)
            {
                p_oAdo.SqlNonQuery(p_oConn,"INSERT INTO rxtemp (scenario_id,rx) VALUES ('" + p_strScenarioId + "'," + strRxArray[x].Trim() + ")");
            }
            p_oAdo.m_strSQL = "INSERT INTO scenario_rx_intensity " +
                             "(scenario_id,rx) " +
                             "SELECT a.scenario_id, a.RX " +
                             "FROM rxtemp a " +
                             "WHERE NOT EXISTS (SELECT b.scenario_id,b.rx " +
                                               "FROM scenario_rx_intensity b " +
                                               "WHERE a.rx=b.rx AND TRIM(UCASE(a.scenario_id))=TRIM(UCASE(b.scenario_id)))";
            p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
            if (p_oAdo.TableExist(p_oConn, "rxtemp"))
                p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE rxtemp");
            //
            //LOAD RXINTENSITY ITEMS
            //
            p_oAdo.m_strSQL = "SELECT * FROM " +
                              Tables.CoreScenarioRuleDefinitions.DefaultScenarioRxIntensityTableName + " " +
                              "WHERE TRIM(UCASE(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "'";

            p_oAdo.SqlQueryReader(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows)
            {
                while (p_oAdo.m_OleDbDataReader.Read())
                {


                    if (p_oAdo.m_OleDbDataReader["rx"] != System.DBNull.Value)
                    {
                        CoreAnalysisScenarioItem.RxIntensityItem oItem = new CoreAnalysisScenarioItem.RxIntensityItem();
                        oItem.Rx = p_oAdo.m_OleDbDataReader["rx"].ToString().Trim();
                        if (p_oAdo.m_OleDbDataReader["rx_intensity"] != System.DBNull.Value)
                        {
                            oItem.RxIntensity = Convert.ToInt32(p_oAdo.m_OleDbDataReader["rx_intensity"]);
                        }
                        else
                        {
                            oItem.RxIntensity = -1;
                        }
                        p_oCoreAnalysisScenarioItem.m_oRxIntensityItem_Collection.Add(oItem);
                    }

                }
            }

            p_oAdo.m_OleDbDataReader.Close();
            

        }
        
        public void LoadProcessorScenarioItems(ado_data_access p_oAdo,
                                         System.Data.OleDb.OleDbConnection p_oConn,
                                         string p_strScenarioId,
                                         CoreAnalysisScenarioItem p_oCoreAnalysisScenarioItem)
        {
            int x;
            Queries oQueries = new Queries();
            ProcessorScenarioTools oTools = new ProcessorScenarioTools();

            //scenario mdb connection
            string strProcessorScenarioMDB =
              frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
              "\\processor\\db\\scenario_processor_rule_definitions.mdb";
            //
            //get a list of all the scenarios
            //
            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(strProcessorScenarioMDB, "", ""));
            string[] strScenarioArray =
                frmMain.g_oUtils.ConvertListToArray(
                    oAdo.CreateCommaDelimitedList(
                       oAdo.m_OleDbConnection,
                       "SELECT scenario_id " +
                       "FROM scenario " +
                       "WHERE scenario_id IS NOT NULL AND " +
                                         "LEN(TRIM(scenario_id)) > 0", ""), ",");
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            if (strScenarioArray == null) return;

            for (x = 0; x <= strScenarioArray.Length - 1; x++)
            {
                //
                //LOAD PROJECT DATATASOURCES INFO
                //
                oQueries.m_oFvs.LoadDatasource = true;
                oQueries.m_oFIAPlot.LoadDatasource = true;
                oQueries.m_oProcessor.LoadDatasource = true;
                oQueries.m_oReference.LoadDatasource = true;
                oQueries.LoadDatasources(true, "processor", strScenarioArray[x]);
                oQueries.m_oDataSource.CreateScenarioRuleDefinitionTableLinks(
                    oQueries.m_strTempDbFile,
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim(),
                    "P");
                oTools.LoadAll(oQueries.m_strTempDbFile, oQueries, strScenarioArray[x], p_oCoreAnalysisScenarioItem.m_oProcessorScenarioItem_Collection);
            }

            if (p_oAdo.TableExist(p_oConn, Tables.CoreScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName))
            {
                p_oAdo.m_strSQL = "SELECT * FROM " + Tables.CoreScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName + " " +
                                "WHERE TRIM(UCASE(scenario_id)) = '" +
                                   p_strScenarioId.Trim().ToUpper() + "';";
                p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);
                string strProcessorScenario = "";
                string strFullDetailsYN = "N";

                if (p_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (p_oAdo.m_OleDbDataReader.Read())
                    {
                        if (p_oAdo.m_OleDbDataReader["processor_scenario_id"] != System.DBNull.Value &&
                            p_oAdo.m_OleDbDataReader["processor_scenario_id"].ToString().Trim().Length > 0)
                        {
                            strProcessorScenario = p_oAdo.m_OleDbDataReader["processor_scenario_id"].ToString().Trim();
                        }
                        if (p_oAdo.m_OleDbDataReader["FullDetailsYN"] != System.DBNull.Value &&
                            p_oAdo.m_OleDbDataReader["FullDetailsYN"].ToString().Trim().Length > 0)
                        {
                            strFullDetailsYN = p_oAdo.m_OleDbDataReader["FullDetailsYN"].ToString().Trim();
                        }
                    }
                }
                p_oAdo.m_OleDbDataReader.Close();
                //
                //find the current selected scenario in the collection
                //
                for (x = 0; x <= p_oCoreAnalysisScenarioItem.m_oProcessorScenarioItem_Collection.Count - 1; x++)
                {
                    if (p_oCoreAnalysisScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).ScenarioId.Trim().ToUpper() ==
                        strProcessorScenario.Trim().ToUpper())
                    {
                        p_oCoreAnalysisScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).Selected = true;
                        p_oCoreAnalysisScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).DisplayFullDetailsYN = strFullDetailsYN;
                    }
                    else
                    {
                        p_oCoreAnalysisScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).Selected = false;
                        p_oCoreAnalysisScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).DisplayFullDetailsYN = strFullDetailsYN;

                    }
                }
            }
            else
            {
                frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioProcessorScenarioSelectTable(oAdo, p_oConn, Tables.CoreScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName);
            }

            
            

        }
        public void LoadPlotFilter(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.CoreAnalysisScenarioItem p_oCoreAnalysisScenarioItem)
        {

            p_oAdo.m_strSQL = "SELECT * FROM scenario_plot_filter  WHERE " +
                " TRIM(UCASE(scenario_id)) = '" + p_strScenarioId.Trim().ToUpper() + "' AND current_yn = 'Y';";
            p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);



            if (p_oAdo.m_intError == 0)
            {

                if (p_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (p_oAdo.m_OleDbDataReader.Read())
                    {
                        if (p_oAdo.m_OleDbDataReader["sql_command"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["sql_command"].ToString().Trim().Length > 0)
                            {
                                p_oCoreAnalysisScenarioItem.PlotTableSQLFilter = p_oAdo.m_OleDbDataReader["sql_command"].ToString().Trim();
                            }
                        }

                    }
                }
                else
                {

                    p_oCoreAnalysisScenarioItem.PlotTableSQLFilter = "SELECT @@PlotTable@@.* FROM @@PlotTable@@ WHERE @@PlotTable@@.plot_accessible_yn='Y'";

                }

                p_oAdo.m_OleDbDataReader.Close();
            }



        }
        public void LoadCondFilter(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.CoreAnalysisScenarioItem p_oCoreAnalysisScenarioItem)
        {
            p_oAdo.m_strSQL = "SELECT * FROM scenario_cond_filter  WHERE " +
                " TRIM(UCASE(scenario_id)) = '" + p_strScenarioId.Trim().ToUpper() + "' AND current_yn = 'Y';";
            p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);



            if (p_oAdo.m_intError == 0)
            {

                if (p_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (p_oAdo.m_OleDbDataReader.Read())
                    {
                        if (p_oAdo.m_OleDbDataReader["sql_command"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["sql_command"].ToString().Trim().Length > 0)
                            {
                                p_oCoreAnalysisScenarioItem.m_oCondTableSQLFilter.SQL =
                                    p_oAdo.m_OleDbDataReader["sql_command"].ToString().Trim();
                            }
                        }

                    }
                }
                else
                {
                    p_oCoreAnalysisScenarioItem.m_oCondTableSQLFilter.SQL =
                        "SELECT * FROM @@CondTable@@";
                }

                p_oAdo.m_OleDbDataReader.Close();
            }

            p_oAdo.m_strSQL = "SELECT * FROM scenario_cond_filter_misc WHERE " +
                             " TRIM(UCASE(scenario_id)) = '" + p_strScenarioId.Trim().ToUpper() + "';";
            p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);



            if (p_oAdo.m_intError == 0)
            {
                //load  wind speed class definitions
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                    if (p_oAdo.m_OleDbDataReader["yard_dist"] != System.DBNull.Value)
                    {
                        if (p_oAdo.m_OleDbDataReader["yard_dist"].ToString().Trim().Length > 0)
                        {
                            p_oCoreAnalysisScenarioItem.m_oCondTableSQLFilter.LowSlopeMaximumYardingDistanceFeet = 
                                Convert.ToString(p_oAdo.m_OleDbDataReader["yard_dist"].ToString().Trim());
                            p_oCoreAnalysisScenarioItem.m_oCondTableSQLFilter.SteepSlopeMaximumYardingDistanceFeet=
                                Convert.ToString(p_oAdo.m_OleDbDataReader["yard_dist2"].ToString().Trim());
                        }
                    }
                }
                p_oAdo.m_OleDbDataReader.Close();
            }
        }

        public void LoadProcessingSites(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.CoreAnalysisScenarioItem p_oCoreAnalysisScenarioItem)
        {
			

			
			this.m_intError = 0;
            this.m_strError = "";


            int x, y;
            //
            //GET ALL THE PSITES LISTED IN THE TRAVEL TIMES TABLE
            //
            /*************************************************************************
			 **get the travel times mdb file,table, and connection strings
			 *************************************************************************/
            string strTravelTimeMDBFile = "", strTravelTimeTableName = "", strTravelTimeConn = "";
            p_oAdo.getScenarioDataSourceConnStringAndTable(ref strTravelTimeMDBFile,
                                            ref strTravelTimeTableName, ref strTravelTimeConn,
                                            "TRAVEL TIMES",
                                            p_strScenarioId,
                                            p_oConn);
            //
            //GET THE PSITE LIST
            //
            System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
            p_oAdo.OpenConnection(strTravelTimeConn,ref oConn);
            string strList = p_oAdo.CreateCommaDelimitedList(oConn, "SELECT DISTINCT CSTR(psite_id) AS psite_id FROM " + strTravelTimeTableName, "");
            string[] strArray = frmMain.g_oUtils.ConvertListToArray(strList, ",");
            p_oAdo.CloseConnection(oConn);
            //
            //INITIALIZE PSITE COLLECTION
            //
            p_oCoreAnalysisScenarioItem.m_oProcessingSiteItem_Collection.Clear();
            //
            //FIRST DELETE THE PSITE'S THAT NO LONGER EXIST
            //
            p_oAdo.m_strSQL = "DELETE FROM scenario_psites WHERE TRIM(UCASE(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "' AND psite_id NOT IN (" + strList + ")";
            p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
            //
            //LOAD UP THE PSITES THAT EXIST
            //
            for (x=0;x<=strArray.Length-1;x++)
            {
            	CoreAnalysisScenarioItem.ProcessingSiteItem oItem = new CoreAnalysisScenarioItem.ProcessingSiteItem();
            	oItem.ProcessingSiteId = strArray[x].Trim();
            	p_oCoreAnalysisScenarioItem.m_oProcessingSiteItem_Collection.Add(oItem);
            	
            }
            /*************************************************************************
			 **get the processing sites mdb file,table, and connection strings
			 *************************************************************************/
            string strPSitesMDBFile = "", strPSitesTableName = "", strPSitesConn = "";
            p_oAdo.getScenarioDataSourceConnStringAndTable(ref strPSitesMDBFile,
                                            ref strPSitesTableName, ref strPSitesConn,
                                            "PROCESSING SITES",
                                            p_strScenarioId,
                                            p_oConn);
            //
            //GET THE PSITE LIST
            //
            oConn = new System.Data.OleDb.OleDbConnection();
            p_oAdo.OpenConnection(strPSitesConn,ref oConn);
            p_oAdo.m_strSQL = "SELECT DISTINCT p.psite_id,p.name,p.trancd,p.trancd_def,p.biocd,p.biocd_def,p.exists_yn " + 
				             "FROM " + strPSitesTableName + " p WHERE p.psite_id IN (" + strList + ")";
            p_oAdo.SqlQueryReader(oConn,p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows==true)
            {
				x=0;
				while (p_oAdo.m_OleDbDataReader.Read())
				{
					if (p_oAdo.m_OleDbDataReader["psite_id"] != System.DBNull.Value)
					{
						for (x=0;x<=p_oCoreAnalysisScenarioItem.m_oProcessingSiteItem_Collection.Count-1;x++)
						{
							if (p_oCoreAnalysisScenarioItem.m_oProcessingSiteItem_Collection.Item(x).ProcessingSiteId.Trim() ==
								Convert.ToString(p_oAdo.m_OleDbDataReader["psite_id"]).Trim())
							{
                              //processing site name
							  if (p_oAdo.m_OleDbDataReader["name"] != System.DBNull.Value)
                              {
                                p_oCoreAnalysisScenarioItem.m_oProcessingSiteItem_Collection.Item(x).ProcessingSiteName=
                                    Convert.ToString(p_oAdo.m_OleDbDataReader["name"]).Trim();
                              }
                              //transportation code
                              if (p_oAdo.m_OleDbDataReader["trancd"] != System.DBNull.Value)
                              {
                                  p_oCoreAnalysisScenarioItem.m_oProcessingSiteItem_Collection.Item(x).TransportationCode =
                                      Convert.ToString(p_oAdo.m_OleDbDataReader["trancd"]).Trim();
                              }
                              //biomass code
                              if (p_oAdo.m_OleDbDataReader["biocd"] != System.DBNull.Value)
                              {
                                  p_oCoreAnalysisScenarioItem.m_oProcessingSiteItem_Collection.Item(x).BiomassCode =
                                      Convert.ToString(p_oAdo.m_OleDbDataReader["biocd"]).Trim();
                              }
                              //exists YN
                              if (p_oAdo.m_OleDbDataReader["exists_yn"] != System.DBNull.Value)
                              {
                                  p_oCoreAnalysisScenarioItem.m_oProcessingSiteItem_Collection.Item(x).ProcessingSiteExistYN =
                                      Convert.ToString(p_oAdo.m_OleDbDataReader["exists_yn"]).Trim();
                              }
                              break;
							}
						}
					}
				}
            }
            p_oAdo.m_OleDbDataReader.Close();
            p_oAdo.CloseConnection(oConn);
            
            //
            //SET THE PREVIOUSLY SELECTED
            //
            p_oAdo.m_strSQL = "SELECT psite_id,trancd,biocd,selected_yn " +
                              "FROM scenario_psites WHERE TRIM(UCASE(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "'";
            p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows == true)
            {
                x = 0;
                while (p_oAdo.m_OleDbDataReader.Read())
                {

                    if (p_oAdo.m_OleDbDataReader["psite_id"] != System.DBNull.Value)
                    {
                        for (x = 0; x <= p_oCoreAnalysisScenarioItem.m_oProcessingSiteItem_Collection.Count - 1; x++)
                        {
                            if (p_oCoreAnalysisScenarioItem.m_oProcessingSiteItem_Collection.Item(x).ProcessingSiteId.Trim() ==
                                Convert.ToString(p_oAdo.m_OleDbDataReader["psite_id"]).Trim())
                            {
                                if (p_oAdo.m_OleDbDataReader["selected_yn"] != System.DBNull.Value)
                                {
                                    if (Convert.ToString(p_oAdo.m_OleDbDataReader["selected_yn"]).Trim() == "Y")
                                    {
                                        p_oCoreAnalysisScenarioItem.m_oProcessingSiteItem_Collection.Item(x).Selected = true;
                                    }
                                    else
                                    {
                                        p_oCoreAnalysisScenarioItem.m_oProcessingSiteItem_Collection.Item(x).Selected = false;
                                    }
                                }
                                //transportation code
                                if (p_oAdo.m_OleDbDataReader["trancd"] != System.DBNull.Value)
                                {
                                    p_oCoreAnalysisScenarioItem.m_oProcessingSiteItem_Collection.Item(x).TransportationCode =
                                        Convert.ToString(p_oAdo.m_OleDbDataReader["trancd"]).Trim();
                                }
                                //biomass code
                                if (p_oAdo.m_OleDbDataReader["biocd"] != System.DBNull.Value)
                                {
                                    p_oCoreAnalysisScenarioItem.m_oProcessingSiteItem_Collection.Item(x).BiomassCode =
                                        Convert.ToString(p_oAdo.m_OleDbDataReader["biocd"]).Trim();
                                }
                                break;
                            }
                        }
                    }
                }
                
            }
            p_oAdo.m_OleDbDataReader.Close();
            

			
		}
        public void LoadLandOwnerGroupFilter(ado_data_access p_oAdo, 
                                         System.Data.OleDb.OleDbConnection p_oConn,
                                         string p_strScenarioId, 
                                         CoreAnalysisScenarioItem p_oCoreAnalysisScenarioItem)
        {
             p_oCoreAnalysisScenarioItem.OwnerGroupCodeList="";
            string strSQL = "SELECT * FROM scenario_land_owner_groups WHERE " +
                " scenario_id = '" + p_strScenarioId + "';";
            p_oAdo.SqlQueryReader(p_oConn, strSQL);


            //insert records into the scenario_rx_intensity table from the master rx table
            if (p_oAdo.m_intError == 0)
            {

                //load step one with wind class speed definitions
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                    if (p_oAdo.m_OleDbDataReader["owngrpcd"] != System.DBNull.Value)
                    {
                        p_oCoreAnalysisScenarioItem.OwnerGroupCodeList =
                            p_oCoreAnalysisScenarioItem.OwnerGroupCodeList +
                            Convert.ToString(p_oAdo.m_OleDbDataReader["owngrpcd"]).Trim() + ",";
                    }
                }
                p_oAdo.m_OleDbDataReader.Close();
                p_oAdo.m_OleDbDataReader = null;
                p_oAdo.m_OleDbCommand = null;
                if (p_oCoreAnalysisScenarioItem.OwnerGroupCodeList.Trim().Length > 0)
                    p_oCoreAnalysisScenarioItem.OwnerGroupCodeList =
                        p_oCoreAnalysisScenarioItem.OwnerGroupCodeList.Substring(0, p_oCoreAnalysisScenarioItem.OwnerGroupCodeList.Length - 1);
            }
        }
        public void LoadTransportationCosts(ado_data_access p_oAdo,
                                         System.Data.OleDb.OleDbConnection p_oConn,
                                         string p_strScenarioId,
                                         CoreAnalysisScenarioItem p_oCoreAnalysisScenarioItem)
        {
            p_oCoreAnalysisScenarioItem.m_oTranCosts.RailChipTransferPerGreenTonPerHour = "";
            p_oCoreAnalysisScenarioItem.m_oTranCosts.RailHaulCostPerGreenTonPerMile = "";
            p_oCoreAnalysisScenarioItem.m_oTranCosts.RailMerchTransferPerGreenTonPerHour = "";
            p_oCoreAnalysisScenarioItem.m_oTranCosts.RoadHaulCostPerGreenTonPerHour="";

            
            string strSQL = "SELECT * FROM scenario_costs WHERE " +
                " TRIM(UCASE(scenario_id)) = '" + p_strScenarioId.Trim().ToUpper() + "';";
            p_oAdo.SqlQueryReader(p_oConn, strSQL);


            //insert records into the scenario_rx_intensity table from the master rx table
            if (p_oAdo.m_intError == 0)
            {

                //load step one with wind class speed definitions
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                    if (p_oAdo.m_OleDbDataReader["road_haul_cost_pgt_per_hour"] != System.DBNull.Value)
                    {
                        p_oCoreAnalysisScenarioItem.m_oTranCosts.RoadHaulCostPerGreenTonPerHour =
                           Convert.ToString(p_oAdo.m_OleDbDataReader["road_haul_cost_pgt_per_hour"]).Trim();
                    }
                    if (p_oAdo.m_OleDbDataReader["rail_haul_cost_pgt_per_mile"] != System.DBNull.Value)
                    {
                        p_oCoreAnalysisScenarioItem.m_oTranCosts.RailHaulCostPerGreenTonPerMile =
                           Convert.ToString(p_oAdo.m_OleDbDataReader["rail_haul_cost_pgt_per_mile"]).Trim();
                    }
                    if (p_oAdo.m_OleDbDataReader["rail_chip_transfer_pgt_per_hour"] != System.DBNull.Value)
                    {
                        p_oCoreAnalysisScenarioItem.m_oTranCosts.RailChipTransferPerGreenTonPerHour =
                           Convert.ToString(p_oAdo.m_OleDbDataReader["rail_chip_transfer_pgt_per_hour"]).Trim();
                    }
                    if (p_oAdo.m_OleDbDataReader["rail_merch_transfer_pgt_per_hour"] != System.DBNull.Value)
                    {
                        p_oCoreAnalysisScenarioItem.m_oTranCosts.RailMerchTransferPerGreenTonPerHour =
                           Convert.ToString(p_oAdo.m_OleDbDataReader["rail_merch_transfer_pgt_per_hour"]).Trim();
                    }

                }
                p_oAdo.m_OleDbDataReader.Close();
               
            }
        }
        public string ScenarioProperties(CoreAnalysisScenarioItem p_oScenarioItem)
        {
            int x,y;
            bool bUseRxIntensity=false;
            string strTranDef="";
            string strBioDef="";
            string strLine = "";
            FIA_Biosum_Manager.ValidateNumericValues oValidate = new ValidateNumericValues();
            strLine = "***CORE ANALYSIS SCENARIO***\r\n\r\n";
            strLine = strLine + "ID\r\n";
            strLine = strLine + "-----\r\n";
            strLine = strLine + p_oScenarioItem.ScenarioId + "\r\n\r\n";
            strLine = strLine + "Description\r\n";
            strLine = strLine + "---------------------\r\n";
            strLine = strLine + p_oScenarioItem.Description + "\r\n\r\n";
            strLine = strLine + "FVS Analysis Cycle 1\r\n";
            strLine = strLine + "-------------------------------\r\n";
           
            strLine = strLine + "//\r\n";
            strLine = strLine + "//EFFECTIVE TREATMENT EVALUATION\r\n";
            strLine = strLine + "//\r\n";
            for (x = 0; x <= p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Count - 1; x++)
            {
                if (p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).RxCycle == "1")
                {
                    for (y = 0; y <= CoreAnalysisScenarioItem.EffectiveVariablesItem.NUMBER_OF_VARIABLES - 1; y++)
                    {
                        strLine = strLine + "-FVS Variable " + Convert.ToString(y + 1) + "-\r\n";
                        strLine = strLine + "Pre-Treatment Variable: " +
                            p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strPreVarArray[y] + "\r\n";
                        strLine = strLine + "Post-Treatment Variable: " +
                            p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strPostVarArray[y] + "\r\n";
                        strLine = strLine + "Treatment Improvement Expression: " +
                            p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strBetterExpr[y] + "\r\n";
                        strLine = strLine + "Treatment Regression Expression: " +
                            p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strWorseExpr[y] + "\r\n";
                        strLine = strLine + "Treatment Effective Expression: " +
                        p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strEffectiveExpr[y] + "\r\n";

                        strLine = strLine + "\r\n\r\n";

                    }
                    strLine = strLine + "Overall Effective Expression: " +
                         p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strOverallEffectiveExpr + "\r\n";

                    if (p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_bOverallEffectiveNetRevEnabled)
                    {
                        strLine = strLine + "Overall Effective Net Revenue Filter: " +
                         p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strOverallEffectiveNetRevOperator + " " +
                         p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strOverallEffectiveNetRevValue + "\r\n";


                    }
                }
            }
            strLine = strLine + "//\r\n";
            strLine = strLine + "//OPTIMIZATION VARIABLE\r\n";
            strLine = strLine + "//\r\n";
            for (x = 0; x <= p_oScenarioItem.m_oOptimizationVariableItem_Collection.Count - 1; x++)
            {
                if (p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).RxCycle == "1" &&
                    p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).bSelected)
                {
                    
                        
                        strLine = strLine + "Optimization Variable Name: " +
                            p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).strOptimizedVariable + "\r\n";
                        strLine = strLine + "FVS Variable Name: " +
                            p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).strFVSVariableName + "\r\n";
                        strLine = strLine + "Value Source: " +
                            p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).strValueSource + "\r\n";
                        strLine = strLine + "Aggregate Search: ";
                        if (p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).strMaxYN=="Y")
                              strLine = strLine + "Maximum\r\n";
                        else 
                              strLine = strLine + "Minimum\r\n";
                        if (p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).bUseFilter)
                        {
                            strLine = strLine + "Net Revenue Filter: ";
                            strLine = strLine + p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).strFilterOperator + " " + 
                                p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).dblFilterValue.ToString().Trim();
                        }
                         
                       

                        strLine = strLine + "\r\n";

                }
            }
            strLine = strLine + "//\r\n";
            strLine = strLine + "//TIE BREAKERS\r\n";
            strLine = strLine + "//\r\n";
            for (x = 0; x <= p_oScenarioItem.m_oTieBreaker_Collection.Count - 1; x++)
            {
                if (p_oScenarioItem.m_oTieBreaker_Collection.Item(x).RxCycle == "1" &&
                    p_oScenarioItem.m_oTieBreaker_Collection.Item(x).bSelected)
                {
                     strLine = strLine + "Method: " +
                       p_oScenarioItem.m_oTieBreaker_Collection.Item(x).strMethod + "\r\n";

                     if (p_oScenarioItem.m_oTieBreaker_Collection.Item(x).strMethod.Trim().ToUpper() ==
                          "TREATMENT INTENSITY")
                     {
                         bUseRxIntensity = true;
                     }
                     else
                     {
                         strLine = strLine + "FVS Variable Name: " +
                             p_oScenarioItem.m_oTieBreaker_Collection.Item(x).strFVSVariableName + "\r\n";
                         strLine = strLine + "Value Source: " +
                             p_oScenarioItem.m_oTieBreaker_Collection.Item(x).strValueSource + "\r\n";
                     }
                     strLine = strLine + "Aggregate Search: ";
                     if (p_oScenarioItem.m_oTieBreaker_Collection.Item(x).strMaxYN == "Y")
                         strLine = strLine + "Maximum\r\n";
                     else
                         strLine = strLine + "Minimum\r\n";

                     
                    strLine = strLine + "\r\n\r\n";

                }
            }
            if (bUseRxIntensity)
            {
                strLine = strLine + "\r\nTreatment Intensity Assignments\r\n";
                strLine = strLine + "------------------------------------\r\n";
                strLine = strLine + "--Rx-- --Intensity--\r\n";
                
                for (x = 0; x <= p_oScenarioItem.m_oRxIntensityItem_Collection.Count - 1; x++)
                {
                    strLine = strLine + String.Format("{0,-3}{1,11}",
                              " " + p_oScenarioItem.m_oRxIntensityItem_Collection.Item(x).Rx,
                              p_oScenarioItem.m_oRxIntensityItem_Collection.Item(x).RxIntensity) + "\r\n";
                }
            }
            strLine = strLine + "\r\n";
            strLine = strLine + "Processing Sites\r\n";
            strLine = strLine + "----------------\r\n";
            strLine = strLine + "--Id-- -------------Name------------- ---------------------Transportation--------------------- ---------Biomass--------\r\n";
            for (x = 0; x <= p_oScenarioItem.m_oProcessingSiteItem_Collection.Count - 1; x++)
            {
                if (p_oScenarioItem.m_oProcessingSiteItem_Collection.Item(x).Selected)
                {
                    strTranDef = p_oScenarioItem.m_oProcessingSiteItem_Collection.Item(x).m_strTranCdDescArray[
                        Convert.ToInt32(p_oScenarioItem.m_oProcessingSiteItem_Collection.Item(x).TransportationCode) - 1, 1];
                    strBioDef = p_oScenarioItem.m_oProcessingSiteItem_Collection.Item(x).m_strBioCdDescArray[
                        Convert.ToInt32(p_oScenarioItem.m_oProcessingSiteItem_Collection.Item(x).BiomassCode) - 1, 1];
                    strLine = strLine + String.Format("{0,-7}{1,-31}{2,-57}{3,-53}",
                                  " " + p_oScenarioItem.m_oProcessingSiteItem_Collection.Item(x).ProcessingSiteId,
                                  p_oScenarioItem.m_oProcessingSiteItem_Collection.Item(x).ProcessingSiteName,
                                  strTranDef, strBioDef) + "\r\n";
                }
            }

            strLine = strLine + "\r\n";
            strLine = strLine + "Transportation Costs\r\n";
            strLine = strLine + "----------------------------\r\n";
            strLine = strLine + "Road Haul Cost Per Green Ton Per Hour: ";
            oValidate.Money = true;
            oValidate.RoundDecimalLength = 2;
            oValidate.ValidateDecimal(p_oScenarioItem.m_oTranCosts.RoadHaulCostPerGreenTonPerHour);
            if (oValidate.m_intError == 0)
                strLine=strLine + oValidate.ReturnValue + "\r\n";
            else
            {
                strLine = strLine + "NA\r\n";
            }
            strLine = strLine + "Rail Haul Cost Per Green Ton Per Mile: ";
            oValidate.ValidateDecimal(p_oScenarioItem.m_oTranCosts.RailHaulCostPerGreenTonPerMile);
            if (oValidate.m_intError == 0)
                strLine = strLine + oValidate.ReturnValue + "\r\n";
            else
            {
                strLine = strLine + "NA\r\n";
            }
            strLine = strLine + "Transfer Merch To Rail Per Green Ton Per Hour: ";
            oValidate.ValidateDecimal(p_oScenarioItem.m_oTranCosts.RailMerchTransferPerGreenTonPerHour);
            if (oValidate.m_intError == 0)
                strLine = strLine + oValidate.ReturnValue + "\r\n";
            else
            {
                strLine = strLine + "NA\r\n";
            }
            strLine = strLine + "Transfer Chips To Rail Per Green Ton Per Hour: ";
            oValidate.ValidateDecimal(p_oScenarioItem.m_oTranCosts.RailChipTransferPerGreenTonPerHour);
            if (oValidate.m_intError == 0)
                strLine = strLine + oValidate.ReturnValue + "\r\n";
            else
            {
                strLine = strLine + "NA\r\n";
            }

            strLine = strLine + "\r\n";
            strLine = strLine + "Land Ownership Groups Filter\r\n";
            strLine = strLine + "----------------------------\r\n";
            strLine = strLine + "--Cd-- --------Description--------\r\n";
            string[] strOwnCdArray = frmMain.g_oUtils.ConvertListToArray(p_oScenarioItem.OwnerGroupCodeList,",");
            if (strOwnCdArray != null)
            {
                for (x = 0; x <= strOwnCdArray.Length - 1; x++)
                {
                    if (strOwnCdArray[x].Trim() == "40" ||
                        strOwnCdArray[x].Trim() == "4")
                    {
                        strLine = strLine + string.Format("{0,-6}{1,-20}", "  " + strOwnCdArray[x], "Private") + "\r\n";
                    }
                    else if (strOwnCdArray[x].Trim() == "30" ||
                        strOwnCdArray[x].Trim() == "3")
                    {
                        strLine = strLine + string.Format("{0,-6}{1,-20}", "  " + strOwnCdArray[x], "State and Local Government") + "\r\n";
                    }
                    else if (strOwnCdArray[x].Trim() == "20" ||
                            strOwnCdArray[x].Trim() == "2")
                    {
                        strLine = strLine + string.Format("{0,-6}{1,-20}", "  " + strOwnCdArray[x], "Federal (Non-USFS)") + "\r\n";
                    }
                    else if (strOwnCdArray[x].Trim() == "10" ||
                            strOwnCdArray[x].Trim() == "1")
                    {
                        strLine = strLine + string.Format("{0,-6}{1,-20}", "  " + strOwnCdArray[x], "U.S. Forest Service") + "\r\n";
                    }


                }
            }
            strLine = strLine + "\r\n";
            strLine = strLine + "Condition Table SQL Filter\r\n";
            strLine = strLine + "----------------------------\r\n";
            strLine = strLine + "SQL: " + p_oScenarioItem.m_oCondTableSQLFilter.SQL + "\r\n";
            strLine = strLine + "Maximum Low Slope Yarding Distance (Feet): " + p_oScenarioItem.m_oCondTableSQLFilter.LowSlopeMaximumYardingDistanceFeet + "\r\n"; 
            strLine = strLine + "Maximum Steep Slope Yarding Distance (Feet): " + p_oScenarioItem.m_oCondTableSQLFilter.SteepSlopeMaximumYardingDistanceFeet + "\r\n";

            strLine = strLine + "\r\n";
            strLine = strLine + "Plot Table SQL Filter\r\n";
            strLine = strLine + "----------------------------\r\n";
            strLine = strLine + "SQL: " + p_oScenarioItem.PlotTableSQLFilter;


            strLine = strLine + "\r\n\r\nEOF";

            return strLine;

        }
    }
}
