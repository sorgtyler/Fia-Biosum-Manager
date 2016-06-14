using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Xml;
using System.Threading;



namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_plot_input.
	/// </summary>
	public class uc_plot_input : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.GroupBox grpboxInvType;
		private System.Windows.Forms.RadioButton rdoFIADB;
		private System.Windows.Forms.RadioButton rdoIDB;
		private System.Windows.Forms.Button btnInvTypeHelp;
		private System.Windows.Forms.Button btnInvTypePrevious;
		private System.Windows.Forms.Button btnInvTypeNext;
		private System.Windows.Forms.Button btnInvTypeCancel;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.TextBox txtPlot;
		private System.Windows.Forms.Button btnPlotBrowse;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.Button btnCondBrowse;
		private System.Windows.Forms.TextBox txtCond;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.Button btnTreeBrowse;
		private System.Windows.Forms.TextBox txtTree;
		private System.Windows.Forms.RadioButton rdoFilterNone;
		private System.Windows.Forms.RadioButton rdoFilterByMenu;
		private System.Windows.Forms.RadioButton rdoFilterByFile;
		private System.Windows.Forms.TextBox txtFilterByFile;
		private System.Windows.Forms.Button btnFilterByFileBrowse;
		private System.Windows.Forms.GroupBox grpboxFilter;
		private System.Windows.Forms.Button btnFilterHelp;
		private System.Windows.Forms.Button btnFilterPrevious;
		private System.Windows.Forms.Button btnFilterNext;
		private System.Windows.Forms.Button btnFilterCancel;
		private System.Windows.Forms.ListView lstFilterByState;
		private System.Windows.Forms.Button btnFilterByStateUnselect;
		private System.Windows.Forms.Button btnFilterByStateSelect;
		private System.Windows.Forms.GroupBox grpboxFilterByPlot;
		private System.Windows.Forms.Button btnFilterByPlotUnselect;
		private System.Windows.Forms.Button btnFilterByPlotSelect;
		private System.Windows.Forms.ListView lstFilterByPlot;
		private System.Windows.Forms.Button btnFilterByPlotHelp;
		private System.Windows.Forms.Button btnFilterByPlotPrevious;
		private System.Windows.Forms.Button btnFilterByPlotNext;
		private System.Windows.Forms.Button btnFilterByPlotCancel;
		private System.Windows.Forms.Button btnFilterByPlotFinish;
		public int m_DialogHt;
		public int m_DialogWd;
		private System.Windows.Forms.Button btnFilterByStateFinish;
		private System.Windows.Forms.Button btnFilterByStateHelp;
		private System.Windows.Forms.Button btnFilterByStatePrevious;
		private System.Windows.Forms.Button btnFilterByStateNext;
		private System.Windows.Forms.Button btnFilterByStateCancel;
		private System.Windows.Forms.Button btnFilterFinish;
		private System.Windows.Forms.Button btnInvTypeFinish;
		private System.Windows.Forms.GroupBox grpboxFilterByState;
		private string m_strPlotTxtInputFile;
		private string m_strCondTxtInputFile;
		private string m_strTreeTxtInputFile;
		private string m_strSiteTreeTxtInputFile;
		private string m_strTreeRegionalBiomassTxtInputFile;
		private string m_strPopEvalTxtInputFile;
		private string m_strLoadedPopEvalTxtInputFile="";
		private string m_strPopEstUnitTxtInputFile;
		private string m_strLoadedPopEstUnitTxtInputFile="";
		private string m_strPpsaTxtInputFile;
		private string m_strLoadedPpsaTxtInputFile="";
		private string m_strPopStratumTxtInputFile;
		private string m_strLoadedPopStratumTxtInputFile="";
		private string m_strCurrentTxtInputFile="";



		private string m_strLoadedPopEvalInputTable="";
		private string m_strLoadedPopEstUnitInputTable="";
		private string m_strLoadedPpsaInputTable="";
		private string m_strLoadedPopStratumInputTable="";
		private string m_strLoadedFiadbInputFile="";
		private string m_strCurrentFiadbInputFile="";
		private string m_strCurrentFiadbTable="";
		private string m_strCurrentBiosumTable="";
		private string m_strCurrentProcess="";



		private string m_strTempMDBFile;
		private string m_strTempMDBFileConn;
		private System.Data.OleDb.OleDbConnection m_connTempMDBFile;
		private FIA_Biosum_Manager.ado_data_access m_ado;
		private System.Data.DataTable m_dtStateCounty;
		private System.Data.DataTable m_dtPlot;
		private int m_intError;
		//private string m_strError;
		private string m_strPlotTable;
		private string m_strCondTable;
		private string m_strTreeTable;
		private string m_strSiteTreeTable;
		private string m_strTreeRegionalBiomassTable;
		private string m_strPopEvalTable;
		private string m_strPopEstUnitTable;
		private string m_strPpsaTable;
		private string m_strPopStratumTable;
        private string m_strBiosumPopStratumAdjustmentFactorsTable;
        private string m_strTreeMacroPlotBreakPointDiaTable;
		private string m_strSQL;
		private FIA_Biosum_Manager.frmTherm m_frmTherm;
		private string m_strPlotIdList="";
		private System.Windows.Forms.GroupBox groupBox7;
		private System.Windows.Forms.CheckBox chkForested;
		private System.Windows.Forms.CheckBox chkNonForested;
		private System.Windows.Forms.RadioButton rdoAccess;
		private System.Windows.Forms.RadioButton rdoText;
		private System.Windows.Forms.GroupBox grpInputDataSourceType;
		private System.Windows.Forms.GroupBox grpboxFIADBTxtInput;
		private System.Windows.Forms.Button btnFIADBTxtInputFinish;
		private System.Windows.Forms.Button btnFIADBTxtInputHelp;
		private System.Windows.Forms.Button btnFIADBTxtInputPrevious;
		private System.Windows.Forms.Button btnFIADBTxtInputNext;
		private System.Windows.Forms.Button btnFIADBTxtInputCancel;
		private bool m_bAllCountiesSelected=true;
		private bool m_bAllPlotsSelected=true;
		private string m_strStateCountySQL;
		private string m_strStateCountyPlotSQL;
		private string m_strIDBInv;
		private string m_strLoadedFIADBEvalId="";
		private string m_strLoadedFIADBRsCd="";
		private string m_strCurrFIADBEvalId="";
		private string m_strCurrFIADBRsCd="";
		private string m_strTableType;
		private System.Windows.Forms.GroupBox groupBox8;
		private System.Windows.Forms.Button btnMDBTreeBrowse;
		private System.Windows.Forms.TextBox txtMDBTree;
		private System.Windows.Forms.GroupBox groupBox9;
		private System.Windows.Forms.Button btnMDBCondBrowse;
		private System.Windows.Forms.TextBox txtMDBCond;
		private System.Windows.Forms.GroupBox groupBox10;
		private System.Windows.Forms.Button btnMDBPlotBrowse;
		private System.Windows.Forms.TextBox txtMDBPlotTable;
		private System.Windows.Forms.TextBox txtMDBCondTable;
		private System.Windows.Forms.TextBox txtMDBTreeTable;
		private System.Windows.Forms.TextBox txtMDBPlot;
		private System.Windows.Forms.Button btnMDBInputHelp;
		private System.Windows.Forms.Button btnMDBInputPrevious;
		private System.Windows.Forms.Button btnMDBInputNext;
		private System.Windows.Forms.Button btnMDBInputCancel;
		private System.Windows.Forms.GroupBox grpboxMDBInput;
		private System.Windows.Forms.Button btnMDBInputFinish;
		private System.Windows.Forms.Button btnIDBInvAppend;
		private System.Windows.Forms.ListView lstIDBInv;
		private System.Windows.Forms.Button btnIDBInvHelp;
		private System.Windows.Forms.Button btnIDBInvPrevious;
		private System.Windows.Forms.Button btnIDBInvNext;
		private System.Windows.Forms.Button btnIDBInvCancel;
		private System.Windows.Forms.GroupBox grpboxIDBInv;
		private System.Threading.Thread thdProcessRecords;
		private int m_intAddedPlotRows=0;
		private int m_intAddedCondRows=0;
		private int m_intAddedTreeRows=0;
		private int m_intAddedSiteTreeRows=0;
		private int m_intAddedTreeRegionalBiomassRows=0;
		private System.Windows.Forms.GroupBox groupBox6;
		private System.Windows.Forms.Button btnTreeRegionalBiomassBrowse;
		private System.Windows.Forms.TextBox txtTreeRegionalBiomass;
		private System.Windows.Forms.GroupBox groupBox11;
		private System.Windows.Forms.Button btnPopEvalBrowse;
		private System.Windows.Forms.TextBox txtPopEval;
		private System.Windows.Forms.GroupBox groupBox12;
		private System.Windows.Forms.Button btnPopStratumBrowse;
		private System.Windows.Forms.TextBox txtPopStratum;
		private System.Windows.Forms.GroupBox groupBox13;
		private System.Windows.Forms.Button btnPpsaBrowse;
		private System.Windows.Forms.TextBox txtPpsa;
		private System.Windows.Forms.GroupBox groupBox14;
		private System.Windows.Forms.Button btnPopEstUnitBrowse;
		private System.Windows.Forms.TextBox txtPopEstUnit;
		private System.Windows.Forms.GroupBox grpboxFIADBInv;
		private System.Windows.Forms.Button btnFIADBInvAppend;
		private System.Windows.Forms.ListView lstFIADBInv;
		private System.Windows.Forms.Button btnFIADBInvHelp;
		private System.Windows.Forms.Button btnFIADBInvPrevious;
		private System.Windows.Forms.Button btnFIADBInvNext;
		private System.Windows.Forms.Button btnFIADBInvCancel;
		private Datasource m_oDatasource=null;

		private bool m_bLoadStateCountyList=true;
		private bool m_bLoadStateCountyPlotList=true;
		private System.Windows.Forms.GroupBox grpboxMDBFiadbInput;
		private System.Windows.Forms.GroupBox groupBox15;
		private System.Windows.Forms.GroupBox groupBox16;
		private System.Windows.Forms.GroupBox groupBox17;
		private System.Windows.Forms.GroupBox groupBox18;
		private System.Windows.Forms.GroupBox groupBox19;
		private System.Windows.Forms.GroupBox groupBox20;
		private System.Windows.Forms.GroupBox groupBox21;
		private System.Windows.Forms.GroupBox groupBox22;
		private System.Windows.Forms.Button btnboxMDBFiadbInputFile;
		private System.Windows.Forms.ComboBox cmbFiadbPlotTable;
		private System.Windows.Forms.ComboBox cmbFiadbCondTable;
		private System.Windows.Forms.GroupBox groupBox23;
		private System.Windows.Forms.ComboBox cmbFiadbTreeTable;
		private System.Windows.Forms.ComboBox cmbFiadbTreeRegionalBiomassTable;
		private System.Windows.Forms.ComboBox cmbFiadbPopEvalTable;
		private System.Windows.Forms.ComboBox cmbFiadbPopEstUnitTable;
		private System.Windows.Forms.ComboBox cmbFiadbPopStratumTable;
		private System.Windows.Forms.ComboBox cmbFiadbPpsaTable;
		private System.Windows.Forms.Button btnMDBFiadbInputFinish;
		private System.Windows.Forms.Button btnMDBFiadbInputHelp;
		private System.Windows.Forms.Button btnMDBFiadbInputPrev;
		private System.Windows.Forms.Button btnMDBFiadbInputNext;
		private System.Windows.Forms.Button btnMDBFiadbInputCancel;
		private System.Windows.Forms.TextBox txtMDBFiadbInputFile;
		private System.Windows.Forms.GroupBox groupBox24;
		private System.Windows.Forms.ComboBox cmbFiadbSiteTreeTable;
		private System.Windows.Forms.GroupBox groupBox25;
		private System.Windows.Forms.TextBox txtMDBSiteTreeTable;
		private System.Windows.Forms.Button btnMDBSiteTreeBrowse;
		private System.Windows.Forms.TextBox txtMDBSiteTree;
		private System.Windows.Forms.GroupBox groupBox26;
		private System.Windows.Forms.Button btnSiteTreeBrowse;
		private System.Windows.Forms.TextBox txtSiteTree;
        private frmDialog _frmDialog = null;
        private Label label2;
        private ComboBox cmbCondPropPercent;
        private Label label1;
		

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
        public frmDialog ReferenceFormDialog
        {
            set { _frmDialog = value; }
            get { return _frmDialog; }
        }
       
		public uc_plot_input()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			Initialize();

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_plot_input));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.grpboxMDBFiadbInput = new System.Windows.Forms.GroupBox();
            this.groupBox24 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbSiteTreeTable = new System.Windows.Forms.ComboBox();
            this.groupBox23 = new System.Windows.Forms.GroupBox();
            this.txtMDBFiadbInputFile = new System.Windows.Forms.TextBox();
            this.btnboxMDBFiadbInputFile = new System.Windows.Forms.Button();
            this.groupBox15 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbPopEstUnitTable = new System.Windows.Forms.ComboBox();
            this.groupBox16 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbPpsaTable = new System.Windows.Forms.ComboBox();
            this.groupBox17 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbPopStratumTable = new System.Windows.Forms.ComboBox();
            this.groupBox18 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbPopEvalTable = new System.Windows.Forms.ComboBox();
            this.groupBox19 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbTreeRegionalBiomassTable = new System.Windows.Forms.ComboBox();
            this.groupBox20 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbTreeTable = new System.Windows.Forms.ComboBox();
            this.groupBox21 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbCondTable = new System.Windows.Forms.ComboBox();
            this.groupBox22 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbPlotTable = new System.Windows.Forms.ComboBox();
            this.btnMDBFiadbInputFinish = new System.Windows.Forms.Button();
            this.btnMDBFiadbInputHelp = new System.Windows.Forms.Button();
            this.btnMDBFiadbInputPrev = new System.Windows.Forms.Button();
            this.btnMDBFiadbInputNext = new System.Windows.Forms.Button();
            this.btnMDBFiadbInputCancel = new System.Windows.Forms.Button();
            this.grpboxFIADBInv = new System.Windows.Forms.GroupBox();
            this.btnFIADBInvAppend = new System.Windows.Forms.Button();
            this.lstFIADBInv = new System.Windows.Forms.ListView();
            this.btnFIADBInvHelp = new System.Windows.Forms.Button();
            this.btnFIADBInvPrevious = new System.Windows.Forms.Button();
            this.btnFIADBInvNext = new System.Windows.Forms.Button();
            this.btnFIADBInvCancel = new System.Windows.Forms.Button();
            this.grpboxIDBInv = new System.Windows.Forms.GroupBox();
            this.btnIDBInvAppend = new System.Windows.Forms.Button();
            this.lstIDBInv = new System.Windows.Forms.ListView();
            this.btnIDBInvHelp = new System.Windows.Forms.Button();
            this.btnIDBInvPrevious = new System.Windows.Forms.Button();
            this.btnIDBInvNext = new System.Windows.Forms.Button();
            this.btnIDBInvCancel = new System.Windows.Forms.Button();
            this.grpboxMDBInput = new System.Windows.Forms.GroupBox();
            this.groupBox25 = new System.Windows.Forms.GroupBox();
            this.txtMDBSiteTreeTable = new System.Windows.Forms.TextBox();
            this.btnMDBSiteTreeBrowse = new System.Windows.Forms.Button();
            this.txtMDBSiteTree = new System.Windows.Forms.TextBox();
            this.btnMDBInputFinish = new System.Windows.Forms.Button();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.txtMDBTreeTable = new System.Windows.Forms.TextBox();
            this.btnMDBTreeBrowse = new System.Windows.Forms.Button();
            this.txtMDBTree = new System.Windows.Forms.TextBox();
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.txtMDBCondTable = new System.Windows.Forms.TextBox();
            this.btnMDBCondBrowse = new System.Windows.Forms.Button();
            this.txtMDBCond = new System.Windows.Forms.TextBox();
            this.groupBox10 = new System.Windows.Forms.GroupBox();
            this.txtMDBPlotTable = new System.Windows.Forms.TextBox();
            this.btnMDBPlotBrowse = new System.Windows.Forms.Button();
            this.txtMDBPlot = new System.Windows.Forms.TextBox();
            this.btnMDBInputHelp = new System.Windows.Forms.Button();
            this.btnMDBInputPrevious = new System.Windows.Forms.Button();
            this.btnMDBInputNext = new System.Windows.Forms.Button();
            this.btnMDBInputCancel = new System.Windows.Forms.Button();
            this.grpboxFilterByState = new System.Windows.Forms.GroupBox();
            this.btnFilterByStateFinish = new System.Windows.Forms.Button();
            this.btnFilterByStateUnselect = new System.Windows.Forms.Button();
            this.btnFilterByStateSelect = new System.Windows.Forms.Button();
            this.lstFilterByState = new System.Windows.Forms.ListView();
            this.btnFilterByStateHelp = new System.Windows.Forms.Button();
            this.btnFilterByStatePrevious = new System.Windows.Forms.Button();
            this.btnFilterByStateNext = new System.Windows.Forms.Button();
            this.btnFilterByStateCancel = new System.Windows.Forms.Button();
            this.grpboxFilter = new System.Windows.Forms.GroupBox();
            this.btnFilterFinish = new System.Windows.Forms.Button();
            this.btnFilterHelp = new System.Windows.Forms.Button();
            this.btnFilterPrevious = new System.Windows.Forms.Button();
            this.btnFilterNext = new System.Windows.Forms.Button();
            this.btnFilterCancel = new System.Windows.Forms.Button();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.chkNonForested = new System.Windows.Forms.CheckBox();
            this.chkForested = new System.Windows.Forms.CheckBox();
            this.btnFilterByFileBrowse = new System.Windows.Forms.Button();
            this.txtFilterByFile = new System.Windows.Forms.TextBox();
            this.rdoFilterByFile = new System.Windows.Forms.RadioButton();
            this.rdoFilterByMenu = new System.Windows.Forms.RadioButton();
            this.rdoFilterNone = new System.Windows.Forms.RadioButton();
            this.grpboxFIADBTxtInput = new System.Windows.Forms.GroupBox();
            this.groupBox26 = new System.Windows.Forms.GroupBox();
            this.btnSiteTreeBrowse = new System.Windows.Forms.Button();
            this.txtSiteTree = new System.Windows.Forms.TextBox();
            this.groupBox14 = new System.Windows.Forms.GroupBox();
            this.btnPopEstUnitBrowse = new System.Windows.Forms.Button();
            this.txtPopEstUnit = new System.Windows.Forms.TextBox();
            this.groupBox13 = new System.Windows.Forms.GroupBox();
            this.btnPpsaBrowse = new System.Windows.Forms.Button();
            this.txtPpsa = new System.Windows.Forms.TextBox();
            this.groupBox12 = new System.Windows.Forms.GroupBox();
            this.btnPopStratumBrowse = new System.Windows.Forms.Button();
            this.txtPopStratum = new System.Windows.Forms.TextBox();
            this.groupBox11 = new System.Windows.Forms.GroupBox();
            this.btnPopEvalBrowse = new System.Windows.Forms.Button();
            this.txtPopEval = new System.Windows.Forms.TextBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.btnTreeRegionalBiomassBrowse = new System.Windows.Forms.Button();
            this.txtTreeRegionalBiomass = new System.Windows.Forms.TextBox();
            this.btnFIADBTxtInputFinish = new System.Windows.Forms.Button();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.btnTreeBrowse = new System.Windows.Forms.Button();
            this.txtTree = new System.Windows.Forms.TextBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.btnCondBrowse = new System.Windows.Forms.Button();
            this.txtCond = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnPlotBrowse = new System.Windows.Forms.Button();
            this.txtPlot = new System.Windows.Forms.TextBox();
            this.btnFIADBTxtInputHelp = new System.Windows.Forms.Button();
            this.btnFIADBTxtInputPrevious = new System.Windows.Forms.Button();
            this.btnFIADBTxtInputNext = new System.Windows.Forms.Button();
            this.btnFIADBTxtInputCancel = new System.Windows.Forms.Button();
            this.grpboxInvType = new System.Windows.Forms.GroupBox();
            this.btnInvTypeFinish = new System.Windows.Forms.Button();
            this.btnInvTypeHelp = new System.Windows.Forms.Button();
            this.btnInvTypePrevious = new System.Windows.Forms.Button();
            this.btnInvTypeNext = new System.Windows.Forms.Button();
            this.btnInvTypeCancel = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.grpInputDataSourceType = new System.Windows.Forms.GroupBox();
            this.rdoText = new System.Windows.Forms.RadioButton();
            this.rdoAccess = new System.Windows.Forms.RadioButton();
            this.rdoFIADB = new System.Windows.Forms.RadioButton();
            this.rdoIDB = new System.Windows.Forms.RadioButton();
            this.lblTitle = new System.Windows.Forms.Label();
            this.grpboxFilterByPlot = new System.Windows.Forms.GroupBox();
            this.btnFilterByPlotFinish = new System.Windows.Forms.Button();
            this.btnFilterByPlotUnselect = new System.Windows.Forms.Button();
            this.btnFilterByPlotSelect = new System.Windows.Forms.Button();
            this.lstFilterByPlot = new System.Windows.Forms.ListView();
            this.btnFilterByPlotHelp = new System.Windows.Forms.Button();
            this.btnFilterByPlotPrevious = new System.Windows.Forms.Button();
            this.btnFilterByPlotNext = new System.Windows.Forms.Button();
            this.btnFilterByPlotCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.cmbCondPropPercent = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.grpboxMDBFiadbInput.SuspendLayout();
            this.groupBox24.SuspendLayout();
            this.groupBox23.SuspendLayout();
            this.groupBox15.SuspendLayout();
            this.groupBox16.SuspendLayout();
            this.groupBox17.SuspendLayout();
            this.groupBox18.SuspendLayout();
            this.groupBox19.SuspendLayout();
            this.groupBox20.SuspendLayout();
            this.groupBox21.SuspendLayout();
            this.groupBox22.SuspendLayout();
            this.grpboxFIADBInv.SuspendLayout();
            this.grpboxIDBInv.SuspendLayout();
            this.grpboxMDBInput.SuspendLayout();
            this.groupBox25.SuspendLayout();
            this.groupBox8.SuspendLayout();
            this.groupBox9.SuspendLayout();
            this.groupBox10.SuspendLayout();
            this.grpboxFilterByState.SuspendLayout();
            this.grpboxFilter.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.grpboxFIADBTxtInput.SuspendLayout();
            this.groupBox26.SuspendLayout();
            this.groupBox14.SuspendLayout();
            this.groupBox13.SuspendLayout();
            this.groupBox12.SuspendLayout();
            this.groupBox11.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.grpboxInvType.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.grpInputDataSourceType.SuspendLayout();
            this.grpboxFilterByPlot.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.grpboxMDBFiadbInput);
            this.groupBox1.Controls.Add(this.grpboxFIADBInv);
            this.groupBox1.Controls.Add(this.grpboxIDBInv);
            this.groupBox1.Controls.Add(this.grpboxMDBInput);
            this.groupBox1.Controls.Add(this.grpboxFilterByState);
            this.groupBox1.Controls.Add(this.grpboxFilter);
            this.groupBox1.Controls.Add(this.grpboxFIADBTxtInput);
            this.groupBox1.Controls.Add(this.grpboxInvType);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Controls.Add(this.grpboxFilterByPlot);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(704, 3500);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // grpboxMDBFiadbInput
            // 
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox24);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox23);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox15);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox16);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox17);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox18);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox19);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox20);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox21);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox22);
            this.grpboxMDBFiadbInput.Controls.Add(this.btnMDBFiadbInputFinish);
            this.grpboxMDBFiadbInput.Controls.Add(this.btnMDBFiadbInputHelp);
            this.grpboxMDBFiadbInput.Controls.Add(this.btnMDBFiadbInputPrev);
            this.grpboxMDBFiadbInput.Controls.Add(this.btnMDBFiadbInputNext);
            this.grpboxMDBFiadbInput.Controls.Add(this.btnMDBFiadbInputCancel);
            this.grpboxMDBFiadbInput.Location = new System.Drawing.Point(16, 3028);
            this.grpboxMDBFiadbInput.Name = "grpboxMDBFiadbInput";
            this.grpboxMDBFiadbInput.Size = new System.Drawing.Size(672, 360);
            this.grpboxMDBFiadbInput.TabIndex = 35;
            this.grpboxMDBFiadbInput.TabStop = false;
            this.grpboxMDBFiadbInput.Text = "FIADB Microsoft Access Database File Input";
            // 
            // groupBox24
            // 
            this.groupBox24.Controls.Add(this.cmbFiadbSiteTreeTable);
            this.groupBox24.Location = new System.Drawing.Point(19, 272);
            this.groupBox24.Name = "groupBox24";
            this.groupBox24.Size = new System.Drawing.Size(312, 45);
            this.groupBox24.TabIndex = 51;
            this.groupBox24.TabStop = false;
            this.groupBox24.Text = "Site Tree ";
            // 
            // cmbFiadbSiteTreeTable
            // 
            this.cmbFiadbSiteTreeTable.Location = new System.Drawing.Point(8, 16);
            this.cmbFiadbSiteTreeTable.Name = "cmbFiadbSiteTreeTable";
            this.cmbFiadbSiteTreeTable.Size = new System.Drawing.Size(296, 21);
            this.cmbFiadbSiteTreeTable.TabIndex = 3;
            // 
            // groupBox23
            // 
            this.groupBox23.Controls.Add(this.txtMDBFiadbInputFile);
            this.groupBox23.Controls.Add(this.btnboxMDBFiadbInputFile);
            this.groupBox23.Location = new System.Drawing.Point(20, 16);
            this.groupBox23.Name = "groupBox23";
            this.groupBox23.Size = new System.Drawing.Size(632, 48);
            this.groupBox23.TabIndex = 50;
            this.groupBox23.TabStop = false;
            this.groupBox23.Text = "MSAccess File";
            // 
            // txtMDBFiadbInputFile
            // 
            this.txtMDBFiadbInputFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBFiadbInputFile.Location = new System.Drawing.Point(8, 13);
            this.txtMDBFiadbInputFile.Name = "txtMDBFiadbInputFile";
            this.txtMDBFiadbInputFile.Size = new System.Drawing.Size(576, 26);
            this.txtMDBFiadbInputFile.TabIndex = 2;
            // 
            // btnboxMDBFiadbInputFile
            // 
            this.btnboxMDBFiadbInputFile.Image = ((System.Drawing.Image)(resources.GetObject("btnboxMDBFiadbInputFile.Image")));
            this.btnboxMDBFiadbInputFile.Location = new System.Drawing.Point(592, 10);
            this.btnboxMDBFiadbInputFile.Name = "btnboxMDBFiadbInputFile";
            this.btnboxMDBFiadbInputFile.Size = new System.Drawing.Size(32, 32);
            this.btnboxMDBFiadbInputFile.TabIndex = 1;
            this.btnboxMDBFiadbInputFile.Click += new System.EventHandler(this.btnboxMDBFiadbInputFile_Click);
            // 
            // groupBox15
            // 
            this.groupBox15.Controls.Add(this.cmbFiadbPopEstUnitTable);
            this.groupBox15.Location = new System.Drawing.Point(340, 118);
            this.groupBox15.Name = "groupBox15";
            this.groupBox15.Size = new System.Drawing.Size(312, 45);
            this.groupBox15.TabIndex = 49;
            this.groupBox15.TabStop = false;
            this.groupBox15.Text = "Population Estimation Unit";
            // 
            // cmbFiadbPopEstUnitTable
            // 
            this.cmbFiadbPopEstUnitTable.Location = new System.Drawing.Point(8, 17);
            this.cmbFiadbPopEstUnitTable.Name = "cmbFiadbPopEstUnitTable";
            this.cmbFiadbPopEstUnitTable.Size = new System.Drawing.Size(296, 21);
            this.cmbFiadbPopEstUnitTable.TabIndex = 4;
            // 
            // groupBox16
            // 
            this.groupBox16.Controls.Add(this.cmbFiadbPpsaTable);
            this.groupBox16.Location = new System.Drawing.Point(340, 220);
            this.groupBox16.Name = "groupBox16";
            this.groupBox16.Size = new System.Drawing.Size(312, 45);
            this.groupBox16.TabIndex = 48;
            this.groupBox16.TabStop = false;
            this.groupBox16.Text = "Population Plot Stratum Assignment";
            // 
            // cmbFiadbPpsaTable
            // 
            this.cmbFiadbPpsaTable.Location = new System.Drawing.Point(8, 16);
            this.cmbFiadbPpsaTable.Name = "cmbFiadbPpsaTable";
            this.cmbFiadbPpsaTable.Size = new System.Drawing.Size(296, 21);
            this.cmbFiadbPpsaTable.TabIndex = 4;
            // 
            // groupBox17
            // 
            this.groupBox17.Controls.Add(this.cmbFiadbPopStratumTable);
            this.groupBox17.Location = new System.Drawing.Point(340, 169);
            this.groupBox17.Name = "groupBox17";
            this.groupBox17.Size = new System.Drawing.Size(312, 45);
            this.groupBox17.TabIndex = 47;
            this.groupBox17.TabStop = false;
            this.groupBox17.Text = "Population Stratum";
            // 
            // cmbFiadbPopStratumTable
            // 
            this.cmbFiadbPopStratumTable.Location = new System.Drawing.Point(8, 17);
            this.cmbFiadbPopStratumTable.Name = "cmbFiadbPopStratumTable";
            this.cmbFiadbPopStratumTable.Size = new System.Drawing.Size(296, 21);
            this.cmbFiadbPopStratumTable.TabIndex = 4;
            // 
            // groupBox18
            // 
            this.groupBox18.Controls.Add(this.cmbFiadbPopEvalTable);
            this.groupBox18.Location = new System.Drawing.Point(340, 71);
            this.groupBox18.Name = "groupBox18";
            this.groupBox18.Size = new System.Drawing.Size(312, 45);
            this.groupBox18.TabIndex = 46;
            this.groupBox18.TabStop = false;
            this.groupBox18.Text = "Population Evaluation";
            // 
            // cmbFiadbPopEvalTable
            // 
            this.cmbFiadbPopEvalTable.Location = new System.Drawing.Point(8, 16);
            this.cmbFiadbPopEvalTable.Name = "cmbFiadbPopEvalTable";
            this.cmbFiadbPopEvalTable.Size = new System.Drawing.Size(296, 21);
            this.cmbFiadbPopEvalTable.TabIndex = 4;
            // 
            // groupBox19
            // 
            this.groupBox19.Controls.Add(this.cmbFiadbTreeRegionalBiomassTable);
            this.groupBox19.Location = new System.Drawing.Point(19, 220);
            this.groupBox19.Name = "groupBox19";
            this.groupBox19.Size = new System.Drawing.Size(312, 45);
            this.groupBox19.TabIndex = 45;
            this.groupBox19.TabStop = false;
            this.groupBox19.Text = "Tree Regional Biomass Data";
            // 
            // cmbFiadbTreeRegionalBiomassTable
            // 
            this.cmbFiadbTreeRegionalBiomassTable.Location = new System.Drawing.Point(8, 16);
            this.cmbFiadbTreeRegionalBiomassTable.Name = "cmbFiadbTreeRegionalBiomassTable";
            this.cmbFiadbTreeRegionalBiomassTable.Size = new System.Drawing.Size(296, 21);
            this.cmbFiadbTreeRegionalBiomassTable.TabIndex = 3;
            // 
            // groupBox20
            // 
            this.groupBox20.Controls.Add(this.cmbFiadbTreeTable);
            this.groupBox20.Location = new System.Drawing.Point(20, 169);
            this.groupBox20.Name = "groupBox20";
            this.groupBox20.Size = new System.Drawing.Size(312, 45);
            this.groupBox20.TabIndex = 44;
            this.groupBox20.TabStop = false;
            this.groupBox20.Text = "Tree Data";
            // 
            // cmbFiadbTreeTable
            // 
            this.cmbFiadbTreeTable.Location = new System.Drawing.Point(8, 14);
            this.cmbFiadbTreeTable.Name = "cmbFiadbTreeTable";
            this.cmbFiadbTreeTable.Size = new System.Drawing.Size(296, 21);
            this.cmbFiadbTreeTable.TabIndex = 2;
            // 
            // groupBox21
            // 
            this.groupBox21.Controls.Add(this.cmbFiadbCondTable);
            this.groupBox21.Location = new System.Drawing.Point(20, 118);
            this.groupBox21.Name = "groupBox21";
            this.groupBox21.Size = new System.Drawing.Size(312, 45);
            this.groupBox21.TabIndex = 43;
            this.groupBox21.TabStop = false;
            this.groupBox21.Text = "Condition Data";
            // 
            // cmbFiadbCondTable
            // 
            this.cmbFiadbCondTable.Location = new System.Drawing.Point(8, 16);
            this.cmbFiadbCondTable.Name = "cmbFiadbCondTable";
            this.cmbFiadbCondTable.Size = new System.Drawing.Size(296, 21);
            this.cmbFiadbCondTable.TabIndex = 1;
            // 
            // groupBox22
            // 
            this.groupBox22.Controls.Add(this.cmbFiadbPlotTable);
            this.groupBox22.Location = new System.Drawing.Point(20, 71);
            this.groupBox22.Name = "groupBox22";
            this.groupBox22.Size = new System.Drawing.Size(312, 45);
            this.groupBox22.TabIndex = 42;
            this.groupBox22.TabStop = false;
            this.groupBox22.Text = "Plot Data";
            // 
            // cmbFiadbPlotTable
            // 
            this.cmbFiadbPlotTable.Location = new System.Drawing.Point(7, 16);
            this.cmbFiadbPlotTable.Name = "cmbFiadbPlotTable";
            this.cmbFiadbPlotTable.Size = new System.Drawing.Size(297, 21);
            this.cmbFiadbPlotTable.TabIndex = 0;
            // 
            // btnMDBFiadbInputFinish
            // 
            this.btnMDBFiadbInputFinish.Enabled = false;
            this.btnMDBFiadbInputFinish.Location = new System.Drawing.Point(584, 326);
            this.btnMDBFiadbInputFinish.Name = "btnMDBFiadbInputFinish";
            this.btnMDBFiadbInputFinish.Size = new System.Drawing.Size(72, 24);
            this.btnMDBFiadbInputFinish.TabIndex = 7;
            this.btnMDBFiadbInputFinish.Text = "Append";
            // 
            // btnMDBFiadbInputHelp
            // 
            this.btnMDBFiadbInputHelp.Location = new System.Drawing.Point(24, 326);
            this.btnMDBFiadbInputHelp.Name = "btnMDBFiadbInputHelp";
            this.btnMDBFiadbInputHelp.Size = new System.Drawing.Size(64, 24);
            this.btnMDBFiadbInputHelp.TabIndex = 3;
            this.btnMDBFiadbInputHelp.Text = "Help";
            // 
            // btnMDBFiadbInputPrev
            // 
            this.btnMDBFiadbInputPrev.Enabled = false;
            this.btnMDBFiadbInputPrev.Location = new System.Drawing.Point(424, 326);
            this.btnMDBFiadbInputPrev.Name = "btnMDBFiadbInputPrev";
            this.btnMDBFiadbInputPrev.Size = new System.Drawing.Size(72, 24);
            this.btnMDBFiadbInputPrev.TabIndex = 5;
            this.btnMDBFiadbInputPrev.Text = "< Previous";
            this.btnMDBFiadbInputPrev.Click += new System.EventHandler(this.btnMDBFiadbInputPrev_Click);
            // 
            // btnMDBFiadbInputNext
            // 
            this.btnMDBFiadbInputNext.Location = new System.Drawing.Point(496, 326);
            this.btnMDBFiadbInputNext.Name = "btnMDBFiadbInputNext";
            this.btnMDBFiadbInputNext.Size = new System.Drawing.Size(72, 24);
            this.btnMDBFiadbInputNext.TabIndex = 6;
            this.btnMDBFiadbInputNext.Text = "Next >";
            this.btnMDBFiadbInputNext.Click += new System.EventHandler(this.btnMDBFiadbInputNext_Click);
            // 
            // btnMDBFiadbInputCancel
            // 
            this.btnMDBFiadbInputCancel.Location = new System.Drawing.Point(336, 326);
            this.btnMDBFiadbInputCancel.Name = "btnMDBFiadbInputCancel";
            this.btnMDBFiadbInputCancel.Size = new System.Drawing.Size(64, 24);
            this.btnMDBFiadbInputCancel.TabIndex = 4;
            this.btnMDBFiadbInputCancel.Text = "Cancel";
            this.btnMDBFiadbInputCancel.Click += new System.EventHandler(this.btnMDBFiadbInputCancel_Click);
            // 
            // grpboxFIADBInv
            // 
            this.grpboxFIADBInv.Controls.Add(this.btnFIADBInvAppend);
            this.grpboxFIADBInv.Controls.Add(this.lstFIADBInv);
            this.grpboxFIADBInv.Controls.Add(this.btnFIADBInvHelp);
            this.grpboxFIADBInv.Controls.Add(this.btnFIADBInvPrevious);
            this.grpboxFIADBInv.Controls.Add(this.btnFIADBInvNext);
            this.grpboxFIADBInv.Controls.Add(this.btnFIADBInvCancel);
            this.grpboxFIADBInv.Location = new System.Drawing.Point(16, 2660);
            this.grpboxFIADBInv.Name = "grpboxFIADBInv";
            this.grpboxFIADBInv.Size = new System.Drawing.Size(672, 360);
            this.grpboxFIADBInv.TabIndex = 34;
            this.grpboxFIADBInv.TabStop = false;
            this.grpboxFIADBInv.Text = "Select FIADB Inventory Evalulation";
            // 
            // btnFIADBInvAppend
            // 
            this.btnFIADBInvAppend.Enabled = false;
            this.btnFIADBInvAppend.Location = new System.Drawing.Point(584, 325);
            this.btnFIADBInvAppend.Name = "btnFIADBInvAppend";
            this.btnFIADBInvAppend.Size = new System.Drawing.Size(72, 24);
            this.btnFIADBInvAppend.TabIndex = 34;
            this.btnFIADBInvAppend.Text = "Append";
            this.btnFIADBInvAppend.Click += new System.EventHandler(this.btnFIADBInvAppend_Click);
            // 
            // lstFIADBInv
            // 
            this.lstFIADBInv.FullRowSelect = true;
            this.lstFIADBInv.GridLines = true;
            this.lstFIADBInv.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstFIADBInv.HideSelection = false;
            this.lstFIADBInv.Location = new System.Drawing.Point(16, 32);
            this.lstFIADBInv.MultiSelect = false;
            this.lstFIADBInv.Name = "lstFIADBInv";
            this.lstFIADBInv.Size = new System.Drawing.Size(640, 280);
            this.lstFIADBInv.TabIndex = 30;
            this.lstFIADBInv.UseCompatibleStateImageBehavior = false;
            this.lstFIADBInv.View = System.Windows.Forms.View.Details;
            this.lstFIADBInv.SelectedIndexChanged += new System.EventHandler(this.lstFIADBInv_SelectedIndexChanged);
            // 
            // btnFIADBInvHelp
            // 
            this.btnFIADBInvHelp.Location = new System.Drawing.Point(16, 325);
            this.btnFIADBInvHelp.Name = "btnFIADBInvHelp";
            this.btnFIADBInvHelp.Size = new System.Drawing.Size(64, 24);
            this.btnFIADBInvHelp.TabIndex = 23;
            this.btnFIADBInvHelp.Text = "Help";
            // 
            // btnFIADBInvPrevious
            // 
            this.btnFIADBInvPrevious.Location = new System.Drawing.Point(424, 325);
            this.btnFIADBInvPrevious.Name = "btnFIADBInvPrevious";
            this.btnFIADBInvPrevious.Size = new System.Drawing.Size(72, 24);
            this.btnFIADBInvPrevious.TabIndex = 22;
            this.btnFIADBInvPrevious.Text = "< Previous";
            this.btnFIADBInvPrevious.Click += new System.EventHandler(this.btnFIADBInvPrevious_Click);
            // 
            // btnFIADBInvNext
            // 
            this.btnFIADBInvNext.Location = new System.Drawing.Point(496, 325);
            this.btnFIADBInvNext.Name = "btnFIADBInvNext";
            this.btnFIADBInvNext.Size = new System.Drawing.Size(72, 24);
            this.btnFIADBInvNext.TabIndex = 21;
            this.btnFIADBInvNext.Text = "Next >";
            this.btnFIADBInvNext.Click += new System.EventHandler(this.btnFIADBInvNext_Click);
            // 
            // btnFIADBInvCancel
            // 
            this.btnFIADBInvCancel.Location = new System.Drawing.Point(336, 325);
            this.btnFIADBInvCancel.Name = "btnFIADBInvCancel";
            this.btnFIADBInvCancel.Size = new System.Drawing.Size(64, 24);
            this.btnFIADBInvCancel.TabIndex = 20;
            this.btnFIADBInvCancel.Text = "Cancel";
            this.btnFIADBInvCancel.Click += new System.EventHandler(this.btnFIADBInvCancel_Click);
            // 
            // grpboxIDBInv
            // 
            this.grpboxIDBInv.Controls.Add(this.btnIDBInvAppend);
            this.grpboxIDBInv.Controls.Add(this.lstIDBInv);
            this.grpboxIDBInv.Controls.Add(this.btnIDBInvHelp);
            this.grpboxIDBInv.Controls.Add(this.btnIDBInvPrevious);
            this.grpboxIDBInv.Controls.Add(this.btnIDBInvNext);
            this.grpboxIDBInv.Controls.Add(this.btnIDBInvCancel);
            this.grpboxIDBInv.Location = new System.Drawing.Point(16, 2284);
            this.grpboxIDBInv.Name = "grpboxIDBInv";
            this.grpboxIDBInv.Size = new System.Drawing.Size(672, 360);
            this.grpboxIDBInv.TabIndex = 33;
            this.grpboxIDBInv.TabStop = false;
            this.grpboxIDBInv.Text = "Select IDB Inventory";
            // 
            // btnIDBInvAppend
            // 
            this.btnIDBInvAppend.Enabled = false;
            this.btnIDBInvAppend.Location = new System.Drawing.Point(584, 326);
            this.btnIDBInvAppend.Name = "btnIDBInvAppend";
            this.btnIDBInvAppend.Size = new System.Drawing.Size(72, 24);
            this.btnIDBInvAppend.TabIndex = 34;
            this.btnIDBInvAppend.Text = "Append";
            this.btnIDBInvAppend.Click += new System.EventHandler(this.btnIDBInvAppend_Click);
            // 
            // lstIDBInv
            // 
            this.lstIDBInv.FullRowSelect = true;
            this.lstIDBInv.GridLines = true;
            this.lstIDBInv.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstIDBInv.HideSelection = false;
            this.lstIDBInv.Location = new System.Drawing.Point(16, 32);
            this.lstIDBInv.MultiSelect = false;
            this.lstIDBInv.Name = "lstIDBInv";
            this.lstIDBInv.Size = new System.Drawing.Size(640, 280);
            this.lstIDBInv.TabIndex = 30;
            this.lstIDBInv.UseCompatibleStateImageBehavior = false;
            this.lstIDBInv.View = System.Windows.Forms.View.Details;
            this.lstIDBInv.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.lstIDBInv_ItemCheck);
            // 
            // btnIDBInvHelp
            // 
            this.btnIDBInvHelp.Location = new System.Drawing.Point(16, 326);
            this.btnIDBInvHelp.Name = "btnIDBInvHelp";
            this.btnIDBInvHelp.Size = new System.Drawing.Size(64, 24);
            this.btnIDBInvHelp.TabIndex = 23;
            this.btnIDBInvHelp.Text = "Help";
            // 
            // btnIDBInvPrevious
            // 
            this.btnIDBInvPrevious.Location = new System.Drawing.Point(424, 326);
            this.btnIDBInvPrevious.Name = "btnIDBInvPrevious";
            this.btnIDBInvPrevious.Size = new System.Drawing.Size(72, 24);
            this.btnIDBInvPrevious.TabIndex = 22;
            this.btnIDBInvPrevious.Text = "< Previous";
            this.btnIDBInvPrevious.Click += new System.EventHandler(this.btnIDBInvPrevious_Click);
            // 
            // btnIDBInvNext
            // 
            this.btnIDBInvNext.Location = new System.Drawing.Point(496, 326);
            this.btnIDBInvNext.Name = "btnIDBInvNext";
            this.btnIDBInvNext.Size = new System.Drawing.Size(72, 24);
            this.btnIDBInvNext.TabIndex = 21;
            this.btnIDBInvNext.Text = "Next >";
            this.btnIDBInvNext.Click += new System.EventHandler(this.btnIDBInvNext_Click);
            // 
            // btnIDBInvCancel
            // 
            this.btnIDBInvCancel.Location = new System.Drawing.Point(336, 326);
            this.btnIDBInvCancel.Name = "btnIDBInvCancel";
            this.btnIDBInvCancel.Size = new System.Drawing.Size(64, 24);
            this.btnIDBInvCancel.TabIndex = 20;
            this.btnIDBInvCancel.Text = "Cancel";
            this.btnIDBInvCancel.Click += new System.EventHandler(this.btnIDBInvCancel_Click);
            // 
            // grpboxMDBInput
            // 
            this.grpboxMDBInput.Controls.Add(this.groupBox25);
            this.grpboxMDBInput.Controls.Add(this.btnMDBInputFinish);
            this.grpboxMDBInput.Controls.Add(this.groupBox8);
            this.grpboxMDBInput.Controls.Add(this.groupBox9);
            this.grpboxMDBInput.Controls.Add(this.groupBox10);
            this.grpboxMDBInput.Controls.Add(this.btnMDBInputHelp);
            this.grpboxMDBInput.Controls.Add(this.btnMDBInputPrevious);
            this.grpboxMDBInput.Controls.Add(this.btnMDBInputNext);
            this.grpboxMDBInput.Controls.Add(this.btnMDBInputCancel);
            this.grpboxMDBInput.Location = new System.Drawing.Point(16, 1916);
            this.grpboxMDBInput.Name = "grpboxMDBInput";
            this.grpboxMDBInput.Size = new System.Drawing.Size(672, 360);
            this.grpboxMDBInput.TabIndex = 0;
            this.grpboxMDBInput.TabStop = false;
            this.grpboxMDBInput.Text = "FIADB Microsoft Access Database File Input";
            // 
            // groupBox25
            // 
            this.groupBox25.Controls.Add(this.txtMDBSiteTreeTable);
            this.groupBox25.Controls.Add(this.btnMDBSiteTreeBrowse);
            this.groupBox25.Controls.Add(this.txtMDBSiteTree);
            this.groupBox25.Location = new System.Drawing.Point(24, 248);
            this.groupBox25.Name = "groupBox25";
            this.groupBox25.Size = new System.Drawing.Size(624, 73);
            this.groupBox25.TabIndex = 31;
            this.groupBox25.TabStop = false;
            this.groupBox25.Text = "Site Tree Data";
            // 
            // txtMDBSiteTreeTable
            // 
            this.txtMDBSiteTreeTable.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBSiteTreeTable.Location = new System.Drawing.Point(408, 27);
            this.txtMDBSiteTreeTable.Name = "txtMDBSiteTreeTable";
            this.txtMDBSiteTreeTable.Size = new System.Drawing.Size(152, 26);
            this.txtMDBSiteTreeTable.TabIndex = 1;
            // 
            // btnMDBSiteTreeBrowse
            // 
            this.btnMDBSiteTreeBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnMDBSiteTreeBrowse.Image")));
            this.btnMDBSiteTreeBrowse.Location = new System.Drawing.Point(573, 23);
            this.btnMDBSiteTreeBrowse.Name = "btnMDBSiteTreeBrowse";
            this.btnMDBSiteTreeBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnMDBSiteTreeBrowse.TabIndex = 2;
            this.btnMDBSiteTreeBrowse.Click += new System.EventHandler(this.btnMDBSiteTreeBrowse_Click);
            // 
            // txtMDBSiteTree
            // 
            this.txtMDBSiteTree.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBSiteTree.Location = new System.Drawing.Point(17, 27);
            this.txtMDBSiteTree.Name = "txtMDBSiteTree";
            this.txtMDBSiteTree.Size = new System.Drawing.Size(383, 26);
            this.txtMDBSiteTree.TabIndex = 0;
            // 
            // btnMDBInputFinish
            // 
            this.btnMDBInputFinish.Enabled = false;
            this.btnMDBInputFinish.Location = new System.Drawing.Point(584, 326);
            this.btnMDBInputFinish.Name = "btnMDBInputFinish";
            this.btnMDBInputFinish.Size = new System.Drawing.Size(72, 24);
            this.btnMDBInputFinish.TabIndex = 7;
            this.btnMDBInputFinish.Text = "Append";
            // 
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.txtMDBTreeTable);
            this.groupBox8.Controls.Add(this.btnMDBTreeBrowse);
            this.groupBox8.Controls.Add(this.txtMDBTree);
            this.groupBox8.Location = new System.Drawing.Point(24, 167);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(624, 73);
            this.groupBox8.TabIndex = 30;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "Tree Data";
            // 
            // txtMDBTreeTable
            // 
            this.txtMDBTreeTable.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBTreeTable.Location = new System.Drawing.Point(408, 27);
            this.txtMDBTreeTable.Name = "txtMDBTreeTable";
            this.txtMDBTreeTable.Size = new System.Drawing.Size(152, 26);
            this.txtMDBTreeTable.TabIndex = 1;
            // 
            // btnMDBTreeBrowse
            // 
            this.btnMDBTreeBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnMDBTreeBrowse.Image")));
            this.btnMDBTreeBrowse.Location = new System.Drawing.Point(573, 23);
            this.btnMDBTreeBrowse.Name = "btnMDBTreeBrowse";
            this.btnMDBTreeBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnMDBTreeBrowse.TabIndex = 2;
            this.btnMDBTreeBrowse.Click += new System.EventHandler(this.btnMDBTreeBrowse_Click);
            // 
            // txtMDBTree
            // 
            this.txtMDBTree.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBTree.Location = new System.Drawing.Point(17, 27);
            this.txtMDBTree.Name = "txtMDBTree";
            this.txtMDBTree.Size = new System.Drawing.Size(383, 26);
            this.txtMDBTree.TabIndex = 0;
            // 
            // groupBox9
            // 
            this.groupBox9.Controls.Add(this.txtMDBCondTable);
            this.groupBox9.Controls.Add(this.btnMDBCondBrowse);
            this.groupBox9.Controls.Add(this.txtMDBCond);
            this.groupBox9.Location = new System.Drawing.Point(24, 95);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Size = new System.Drawing.Size(624, 65);
            this.groupBox9.TabIndex = 29;
            this.groupBox9.TabStop = false;
            this.groupBox9.Text = "Condition Data";
            // 
            // txtMDBCondTable
            // 
            this.txtMDBCondTable.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBCondTable.Location = new System.Drawing.Point(408, 24);
            this.txtMDBCondTable.Name = "txtMDBCondTable";
            this.txtMDBCondTable.Size = new System.Drawing.Size(152, 26);
            this.txtMDBCondTable.TabIndex = 1;
            // 
            // btnMDBCondBrowse
            // 
            this.btnMDBCondBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnMDBCondBrowse.Image")));
            this.btnMDBCondBrowse.Location = new System.Drawing.Point(573, 18);
            this.btnMDBCondBrowse.Name = "btnMDBCondBrowse";
            this.btnMDBCondBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnMDBCondBrowse.TabIndex = 2;
            this.btnMDBCondBrowse.Click += new System.EventHandler(this.btnMDBCondBrowse_Click);
            // 
            // txtMDBCond
            // 
            this.txtMDBCond.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBCond.Location = new System.Drawing.Point(17, 24);
            this.txtMDBCond.Name = "txtMDBCond";
            this.txtMDBCond.Size = new System.Drawing.Size(383, 26);
            this.txtMDBCond.TabIndex = 0;
            // 
            // groupBox10
            // 
            this.groupBox10.Controls.Add(this.txtMDBPlotTable);
            this.groupBox10.Controls.Add(this.btnMDBPlotBrowse);
            this.groupBox10.Controls.Add(this.txtMDBPlot);
            this.groupBox10.Location = new System.Drawing.Point(24, 22);
            this.groupBox10.Name = "groupBox10";
            this.groupBox10.Size = new System.Drawing.Size(624, 66);
            this.groupBox10.TabIndex = 28;
            this.groupBox10.TabStop = false;
            this.groupBox10.Text = "Plot Data";
            // 
            // txtMDBPlotTable
            // 
            this.txtMDBPlotTable.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBPlotTable.Location = new System.Drawing.Point(408, 26);
            this.txtMDBPlotTable.Name = "txtMDBPlotTable";
            this.txtMDBPlotTable.Size = new System.Drawing.Size(152, 26);
            this.txtMDBPlotTable.TabIndex = 1;
            // 
            // btnMDBPlotBrowse
            // 
            this.btnMDBPlotBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnMDBPlotBrowse.Image")));
            this.btnMDBPlotBrowse.Location = new System.Drawing.Point(573, 22);
            this.btnMDBPlotBrowse.Name = "btnMDBPlotBrowse";
            this.btnMDBPlotBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnMDBPlotBrowse.TabIndex = 2;
            this.btnMDBPlotBrowse.Click += new System.EventHandler(this.btnMDBPlotBrowse_Click);
            // 
            // txtMDBPlot
            // 
            this.txtMDBPlot.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBPlot.Location = new System.Drawing.Point(17, 26);
            this.txtMDBPlot.Name = "txtMDBPlot";
            this.txtMDBPlot.Size = new System.Drawing.Size(383, 26);
            this.txtMDBPlot.TabIndex = 0;
            // 
            // btnMDBInputHelp
            // 
            this.btnMDBInputHelp.Location = new System.Drawing.Point(24, 326);
            this.btnMDBInputHelp.Name = "btnMDBInputHelp";
            this.btnMDBInputHelp.Size = new System.Drawing.Size(64, 24);
            this.btnMDBInputHelp.TabIndex = 3;
            this.btnMDBInputHelp.Text = "Help";
            // 
            // btnMDBInputPrevious
            // 
            this.btnMDBInputPrevious.Enabled = false;
            this.btnMDBInputPrevious.Location = new System.Drawing.Point(424, 326);
            this.btnMDBInputPrevious.Name = "btnMDBInputPrevious";
            this.btnMDBInputPrevious.Size = new System.Drawing.Size(72, 24);
            this.btnMDBInputPrevious.TabIndex = 5;
            this.btnMDBInputPrevious.Text = "< Previous";
            this.btnMDBInputPrevious.Click += new System.EventHandler(this.btnMDBInputPrevious_Click);
            // 
            // btnMDBInputNext
            // 
            this.btnMDBInputNext.Location = new System.Drawing.Point(496, 326);
            this.btnMDBInputNext.Name = "btnMDBInputNext";
            this.btnMDBInputNext.Size = new System.Drawing.Size(72, 24);
            this.btnMDBInputNext.TabIndex = 6;
            this.btnMDBInputNext.Text = "Next >";
            this.btnMDBInputNext.Click += new System.EventHandler(this.btnMDBInputNext_Click);
            // 
            // btnMDBInputCancel
            // 
            this.btnMDBInputCancel.Location = new System.Drawing.Point(336, 326);
            this.btnMDBInputCancel.Name = "btnMDBInputCancel";
            this.btnMDBInputCancel.Size = new System.Drawing.Size(64, 24);
            this.btnMDBInputCancel.TabIndex = 4;
            this.btnMDBInputCancel.Text = "Cancel";
            this.btnMDBInputCancel.Click += new System.EventHandler(this.btnMDBInputCancel_Click);
            // 
            // grpboxFilterByState
            // 
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateFinish);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateUnselect);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateSelect);
            this.grpboxFilterByState.Controls.Add(this.lstFilterByState);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateHelp);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStatePrevious);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateNext);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateCancel);
            this.grpboxFilterByState.Location = new System.Drawing.Point(16, 1164);
            this.grpboxFilterByState.Name = "grpboxFilterByState";
            this.grpboxFilterByState.Size = new System.Drawing.Size(672, 360);
            this.grpboxFilterByState.TabIndex = 31;
            this.grpboxFilterByState.TabStop = false;
            this.grpboxFilterByState.Text = "Filter By State And County";
            // 
            // btnFilterByStateFinish
            // 
            this.btnFilterByStateFinish.Location = new System.Drawing.Point(584, 325);
            this.btnFilterByStateFinish.Name = "btnFilterByStateFinish";
            this.btnFilterByStateFinish.Size = new System.Drawing.Size(72, 24);
            this.btnFilterByStateFinish.TabIndex = 34;
            this.btnFilterByStateFinish.Text = "Append";
            this.btnFilterByStateFinish.Click += new System.EventHandler(this.btnFilterByStateFinish_Click);
            // 
            // btnFilterByStateUnselect
            // 
            this.btnFilterByStateUnselect.Location = new System.Drawing.Point(560, 182);
            this.btnFilterByStateUnselect.Name = "btnFilterByStateUnselect";
            this.btnFilterByStateUnselect.Size = new System.Drawing.Size(88, 48);
            this.btnFilterByStateUnselect.TabIndex = 32;
            this.btnFilterByStateUnselect.Text = "Clear All";
            this.btnFilterByStateUnselect.Click += new System.EventHandler(this.btnFilterByStateUnselect_Click);
            // 
            // btnFilterByStateSelect
            // 
            this.btnFilterByStateSelect.Location = new System.Drawing.Point(560, 118);
            this.btnFilterByStateSelect.Name = "btnFilterByStateSelect";
            this.btnFilterByStateSelect.Size = new System.Drawing.Size(88, 48);
            this.btnFilterByStateSelect.TabIndex = 31;
            this.btnFilterByStateSelect.Text = "Select All";
            this.btnFilterByStateSelect.Click += new System.EventHandler(this.btnFilterByStateSelect_Click);
            // 
            // lstFilterByState
            // 
            this.lstFilterByState.CheckBoxes = true;
            this.lstFilterByState.FullRowSelect = true;
            this.lstFilterByState.GridLines = true;
            this.lstFilterByState.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstFilterByState.HideSelection = false;
            this.lstFilterByState.Location = new System.Drawing.Point(136, 32);
            this.lstFilterByState.Name = "lstFilterByState";
            this.lstFilterByState.Size = new System.Drawing.Size(400, 280);
            this.lstFilterByState.TabIndex = 30;
            this.lstFilterByState.UseCompatibleStateImageBehavior = false;
            this.lstFilterByState.View = System.Windows.Forms.View.Details;
            this.lstFilterByState.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.lstFilterByState_ItemCheck);
            // 
            // btnFilterByStateHelp
            // 
            this.btnFilterByStateHelp.Location = new System.Drawing.Point(16, 325);
            this.btnFilterByStateHelp.Name = "btnFilterByStateHelp";
            this.btnFilterByStateHelp.Size = new System.Drawing.Size(64, 24);
            this.btnFilterByStateHelp.TabIndex = 23;
            this.btnFilterByStateHelp.Text = "Help";
            // 
            // btnFilterByStatePrevious
            // 
            this.btnFilterByStatePrevious.Location = new System.Drawing.Point(424, 325);
            this.btnFilterByStatePrevious.Name = "btnFilterByStatePrevious";
            this.btnFilterByStatePrevious.Size = new System.Drawing.Size(72, 24);
            this.btnFilterByStatePrevious.TabIndex = 22;
            this.btnFilterByStatePrevious.Text = "< Previous";
            this.btnFilterByStatePrevious.Click += new System.EventHandler(this.btnFilterByStatePrevious_Click);
            // 
            // btnFilterByStateNext
            // 
            this.btnFilterByStateNext.Location = new System.Drawing.Point(496, 325);
            this.btnFilterByStateNext.Name = "btnFilterByStateNext";
            this.btnFilterByStateNext.Size = new System.Drawing.Size(72, 24);
            this.btnFilterByStateNext.TabIndex = 21;
            this.btnFilterByStateNext.Text = "Next >";
            this.btnFilterByStateNext.Click += new System.EventHandler(this.btnFilterByStateNext_Click);
            // 
            // btnFilterByStateCancel
            // 
            this.btnFilterByStateCancel.Location = new System.Drawing.Point(336, 325);
            this.btnFilterByStateCancel.Name = "btnFilterByStateCancel";
            this.btnFilterByStateCancel.Size = new System.Drawing.Size(64, 24);
            this.btnFilterByStateCancel.TabIndex = 20;
            this.btnFilterByStateCancel.Text = "Cancel";
            this.btnFilterByStateCancel.Click += new System.EventHandler(this.btnFilterByStateCancel_Click);
            // 
            // grpboxFilter
            // 
            this.grpboxFilter.Controls.Add(this.btnFilterFinish);
            this.grpboxFilter.Controls.Add(this.btnFilterHelp);
            this.grpboxFilter.Controls.Add(this.btnFilterPrevious);
            this.grpboxFilter.Controls.Add(this.btnFilterNext);
            this.grpboxFilter.Controls.Add(this.btnFilterCancel);
            this.grpboxFilter.Controls.Add(this.groupBox7);
            this.grpboxFilter.Location = new System.Drawing.Point(16, 796);
            this.grpboxFilter.Name = "grpboxFilter";
            this.grpboxFilter.Size = new System.Drawing.Size(672, 360);
            this.grpboxFilter.TabIndex = 30;
            this.grpboxFilter.TabStop = false;
            this.grpboxFilter.Text = "Filter Options";
            // 
            // btnFilterFinish
            // 
            this.btnFilterFinish.Enabled = false;
            this.btnFilterFinish.Location = new System.Drawing.Point(584, 326);
            this.btnFilterFinish.Name = "btnFilterFinish";
            this.btnFilterFinish.Size = new System.Drawing.Size(72, 24);
            this.btnFilterFinish.TabIndex = 5;
            this.btnFilterFinish.Text = "Append";
            this.btnFilterFinish.Click += new System.EventHandler(this.btnFilterFinish_Click);
            // 
            // btnFilterHelp
            // 
            this.btnFilterHelp.Location = new System.Drawing.Point(16, 326);
            this.btnFilterHelp.Name = "btnFilterHelp";
            this.btnFilterHelp.Size = new System.Drawing.Size(64, 24);
            this.btnFilterHelp.TabIndex = 2;
            this.btnFilterHelp.Text = "Help";
            // 
            // btnFilterPrevious
            // 
            this.btnFilterPrevious.Location = new System.Drawing.Point(424, 326);
            this.btnFilterPrevious.Name = "btnFilterPrevious";
            this.btnFilterPrevious.Size = new System.Drawing.Size(72, 24);
            this.btnFilterPrevious.TabIndex = 4;
            this.btnFilterPrevious.Text = "< Previous";
            this.btnFilterPrevious.Click += new System.EventHandler(this.btnFilterPrevious_Click);
            // 
            // btnFilterNext
            // 
            this.btnFilterNext.Enabled = false;
            this.btnFilterNext.Location = new System.Drawing.Point(496, 326);
            this.btnFilterNext.Name = "btnFilterNext";
            this.btnFilterNext.Size = new System.Drawing.Size(72, 24);
            this.btnFilterNext.TabIndex = 5;
            this.btnFilterNext.Text = "Next >";
            this.btnFilterNext.Click += new System.EventHandler(this.btnFilterNext_Click);
            // 
            // btnFilterCancel
            // 
            this.btnFilterCancel.Location = new System.Drawing.Point(336, 326);
            this.btnFilterCancel.Name = "btnFilterCancel";
            this.btnFilterCancel.Size = new System.Drawing.Size(64, 24);
            this.btnFilterCancel.TabIndex = 3;
            this.btnFilterCancel.Text = "Cancel";
            this.btnFilterCancel.Click += new System.EventHandler(this.btnFilterCancel_Click);
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.label2);
            this.groupBox7.Controls.Add(this.cmbCondPropPercent);
            this.groupBox7.Controls.Add(this.label1);
            this.groupBox7.Controls.Add(this.chkNonForested);
            this.groupBox7.Controls.Add(this.chkForested);
            this.groupBox7.Controls.Add(this.btnFilterByFileBrowse);
            this.groupBox7.Controls.Add(this.txtFilterByFile);
            this.groupBox7.Controls.Add(this.rdoFilterByFile);
            this.groupBox7.Controls.Add(this.rdoFilterByMenu);
            this.groupBox7.Controls.Add(this.rdoFilterNone);
            this.groupBox7.Location = new System.Drawing.Point(85, 64);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(519, 249);
            this.groupBox7.TabIndex = 1;
            this.groupBox7.TabStop = false;
            // 
            // chkNonForested
            // 
            this.chkNonForested.Checked = true;
            this.chkNonForested.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkNonForested.Location = new System.Drawing.Point(128, 96);
            this.chkNonForested.Name = "chkNonForested";
            this.chkNonForested.Size = new System.Drawing.Size(112, 16);
            this.chkNonForested.TabIndex = 6;
            this.chkNonForested.Text = "Non Forested";
            // 
            // chkForested
            // 
            this.chkForested.Checked = true;
            this.chkForested.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkForested.Location = new System.Drawing.Point(56, 96);
            this.chkForested.Name = "chkForested";
            this.chkForested.Size = new System.Drawing.Size(72, 16);
            this.chkForested.TabIndex = 5;
            this.chkForested.Text = "Forested";
            // 
            // btnFilterByFileBrowse
            // 
            this.btnFilterByFileBrowse.Enabled = false;
            this.btnFilterByFileBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnFilterByFileBrowse.Image")));
            this.btnFilterByFileBrowse.Location = new System.Drawing.Point(408, 168);
            this.btnFilterByFileBrowse.Name = "btnFilterByFileBrowse";
            this.btnFilterByFileBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnFilterByFileBrowse.TabIndex = 4;
            this.btnFilterByFileBrowse.Click += new System.EventHandler(this.btnFilterByFileBrowse_Click);
            // 
            // txtFilterByFile
            // 
            this.txtFilterByFile.Enabled = false;
            this.txtFilterByFile.Location = new System.Drawing.Point(64, 173);
            this.txtFilterByFile.Name = "txtFilterByFile";
            this.txtFilterByFile.Size = new System.Drawing.Size(328, 20);
            this.txtFilterByFile.TabIndex = 3;
            // 
            // rdoFilterByFile
            // 
            this.rdoFilterByFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoFilterByFile.Location = new System.Drawing.Point(40, 129);
            this.rdoFilterByFile.Name = "rdoFilterByFile";
            this.rdoFilterByFile.Size = new System.Drawing.Size(400, 32);
            this.rdoFilterByFile.TabIndex = 2;
            this.rdoFilterByFile.Text = "Filter By File (Text File Containing Plot_CN numbers)";
            this.rdoFilterByFile.Click += new System.EventHandler(this.rdoFilterByFile_Click);
            // 
            // rdoFilterByMenu
            // 
            this.rdoFilterByMenu.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoFilterByMenu.Location = new System.Drawing.Point(40, 59);
            this.rdoFilterByMenu.Name = "rdoFilterByMenu";
            this.rdoFilterByMenu.Size = new System.Drawing.Size(400, 32);
            this.rdoFilterByMenu.TabIndex = 1;
            this.rdoFilterByMenu.Text = "Filter Plots By Menu Selection (State, County, And Plot)";
            this.rdoFilterByMenu.Click += new System.EventHandler(this.rdoFilterByMenu_Click);
            // 
            // rdoFilterNone
            // 
            this.rdoFilterNone.Checked = true;
            this.rdoFilterNone.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoFilterNone.Location = new System.Drawing.Point(40, 28);
            this.rdoFilterNone.Name = "rdoFilterNone";
            this.rdoFilterNone.Size = new System.Drawing.Size(400, 32);
            this.rdoFilterNone.TabIndex = 0;
            this.rdoFilterNone.TabStop = true;
            this.rdoFilterNone.Text = "Input All Plots";
            this.rdoFilterNone.Click += new System.EventHandler(this.rdoFilterNone_Click);
            // 
            // grpboxFIADBTxtInput
            // 
            this.grpboxFIADBTxtInput.Controls.Add(this.groupBox26);
            this.grpboxFIADBTxtInput.Controls.Add(this.groupBox14);
            this.grpboxFIADBTxtInput.Controls.Add(this.groupBox13);
            this.grpboxFIADBTxtInput.Controls.Add(this.groupBox12);
            this.grpboxFIADBTxtInput.Controls.Add(this.groupBox11);
            this.grpboxFIADBTxtInput.Controls.Add(this.groupBox6);
            this.grpboxFIADBTxtInput.Controls.Add(this.btnFIADBTxtInputFinish);
            this.grpboxFIADBTxtInput.Controls.Add(this.groupBox5);
            this.grpboxFIADBTxtInput.Controls.Add(this.groupBox4);
            this.grpboxFIADBTxtInput.Controls.Add(this.groupBox2);
            this.grpboxFIADBTxtInput.Controls.Add(this.btnFIADBTxtInputHelp);
            this.grpboxFIADBTxtInput.Controls.Add(this.btnFIADBTxtInputPrevious);
            this.grpboxFIADBTxtInput.Controls.Add(this.btnFIADBTxtInputNext);
            this.grpboxFIADBTxtInput.Controls.Add(this.btnFIADBTxtInputCancel);
            this.grpboxFIADBTxtInput.Location = new System.Drawing.Point(16, 424);
            this.grpboxFIADBTxtInput.Name = "grpboxFIADBTxtInput";
            this.grpboxFIADBTxtInput.Size = new System.Drawing.Size(672, 360);
            this.grpboxFIADBTxtInput.TabIndex = 29;
            this.grpboxFIADBTxtInput.TabStop = false;
            this.grpboxFIADBTxtInput.Text = "FIADB Text Input";
            // 
            // groupBox26
            // 
            this.groupBox26.Controls.Add(this.btnSiteTreeBrowse);
            this.groupBox26.Controls.Add(this.txtSiteTree);
            this.groupBox26.Location = new System.Drawing.Point(24, 269);
            this.groupBox26.Name = "groupBox26";
            this.groupBox26.Size = new System.Drawing.Size(313, 56);
            this.groupBox26.TabIndex = 42;
            this.groupBox26.TabStop = false;
            this.groupBox26.Text = "Site Tree Data";
            // 
            // btnSiteTreeBrowse
            // 
            this.btnSiteTreeBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnSiteTreeBrowse.Image")));
            this.btnSiteTreeBrowse.Location = new System.Drawing.Point(272, 16);
            this.btnSiteTreeBrowse.Name = "btnSiteTreeBrowse";
            this.btnSiteTreeBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnSiteTreeBrowse.TabIndex = 1;
            this.btnSiteTreeBrowse.Click += new System.EventHandler(this.btnSiteTreeBrowse_Click);
            // 
            // txtSiteTree
            // 
            this.txtSiteTree.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSiteTree.Location = new System.Drawing.Point(17, 19);
            this.txtSiteTree.Name = "txtSiteTree";
            this.txtSiteTree.Size = new System.Drawing.Size(247, 26);
            this.txtSiteTree.TabIndex = 0;
            // 
            // groupBox14
            // 
            this.groupBox14.Controls.Add(this.btnPopEstUnitBrowse);
            this.groupBox14.Controls.Add(this.txtPopEstUnit);
            this.groupBox14.Location = new System.Drawing.Point(344, 81);
            this.groupBox14.Name = "groupBox14";
            this.groupBox14.Size = new System.Drawing.Size(313, 56);
            this.groupBox14.TabIndex = 41;
            this.groupBox14.TabStop = false;
            this.groupBox14.Text = "Population Estimation Unit";
            // 
            // btnPopEstUnitBrowse
            // 
            this.btnPopEstUnitBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnPopEstUnitBrowse.Image")));
            this.btnPopEstUnitBrowse.Location = new System.Drawing.Point(272, 16);
            this.btnPopEstUnitBrowse.Name = "btnPopEstUnitBrowse";
            this.btnPopEstUnitBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnPopEstUnitBrowse.TabIndex = 1;
            this.btnPopEstUnitBrowse.Click += new System.EventHandler(this.btnPopEstUnitBrowse_Click);
            // 
            // txtPopEstUnit
            // 
            this.txtPopEstUnit.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPopEstUnit.Location = new System.Drawing.Point(17, 19);
            this.txtPopEstUnit.Name = "txtPopEstUnit";
            this.txtPopEstUnit.Size = new System.Drawing.Size(247, 26);
            this.txtPopEstUnit.TabIndex = 0;
            // 
            // groupBox13
            // 
            this.groupBox13.Controls.Add(this.btnPpsaBrowse);
            this.groupBox13.Controls.Add(this.txtPpsa);
            this.groupBox13.Location = new System.Drawing.Point(344, 207);
            this.groupBox13.Name = "groupBox13";
            this.groupBox13.Size = new System.Drawing.Size(313, 56);
            this.groupBox13.TabIndex = 40;
            this.groupBox13.TabStop = false;
            this.groupBox13.Text = "Population Plot Stratum Assignment";
            // 
            // btnPpsaBrowse
            // 
            this.btnPpsaBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnPpsaBrowse.Image")));
            this.btnPpsaBrowse.Location = new System.Drawing.Point(272, 16);
            this.btnPpsaBrowse.Name = "btnPpsaBrowse";
            this.btnPpsaBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnPpsaBrowse.TabIndex = 1;
            this.btnPpsaBrowse.Click += new System.EventHandler(this.btnPpsaBrowse_Click);
            // 
            // txtPpsa
            // 
            this.txtPpsa.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPpsa.Location = new System.Drawing.Point(17, 19);
            this.txtPpsa.Name = "txtPpsa";
            this.txtPpsa.Size = new System.Drawing.Size(247, 26);
            this.txtPpsa.TabIndex = 0;
            // 
            // groupBox12
            // 
            this.groupBox12.Controls.Add(this.btnPopStratumBrowse);
            this.groupBox12.Controls.Add(this.txtPopStratum);
            this.groupBox12.Location = new System.Drawing.Point(344, 143);
            this.groupBox12.Name = "groupBox12";
            this.groupBox12.Size = new System.Drawing.Size(313, 56);
            this.groupBox12.TabIndex = 39;
            this.groupBox12.TabStop = false;
            this.groupBox12.Text = "Population Stratum";
            // 
            // btnPopStratumBrowse
            // 
            this.btnPopStratumBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnPopStratumBrowse.Image")));
            this.btnPopStratumBrowse.Location = new System.Drawing.Point(272, 16);
            this.btnPopStratumBrowse.Name = "btnPopStratumBrowse";
            this.btnPopStratumBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnPopStratumBrowse.TabIndex = 1;
            this.btnPopStratumBrowse.Click += new System.EventHandler(this.btnPopStratumBrowse_Click);
            // 
            // txtPopStratum
            // 
            this.txtPopStratum.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPopStratum.Location = new System.Drawing.Point(17, 19);
            this.txtPopStratum.Name = "txtPopStratum";
            this.txtPopStratum.Size = new System.Drawing.Size(247, 26);
            this.txtPopStratum.TabIndex = 0;
            // 
            // groupBox11
            // 
            this.groupBox11.Controls.Add(this.btnPopEvalBrowse);
            this.groupBox11.Controls.Add(this.txtPopEval);
            this.groupBox11.Location = new System.Drawing.Point(344, 19);
            this.groupBox11.Name = "groupBox11";
            this.groupBox11.Size = new System.Drawing.Size(313, 56);
            this.groupBox11.TabIndex = 38;
            this.groupBox11.TabStop = false;
            this.groupBox11.Text = "Population Evaluation";
            // 
            // btnPopEvalBrowse
            // 
            this.btnPopEvalBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnPopEvalBrowse.Image")));
            this.btnPopEvalBrowse.Location = new System.Drawing.Point(272, 16);
            this.btnPopEvalBrowse.Name = "btnPopEvalBrowse";
            this.btnPopEvalBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnPopEvalBrowse.TabIndex = 1;
            this.btnPopEvalBrowse.Click += new System.EventHandler(this.btnPopEvalBrowse_Click);
            // 
            // txtPopEval
            // 
            this.txtPopEval.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPopEval.Location = new System.Drawing.Point(17, 19);
            this.txtPopEval.Name = "txtPopEval";
            this.txtPopEval.Size = new System.Drawing.Size(247, 26);
            this.txtPopEval.TabIndex = 0;
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.btnTreeRegionalBiomassBrowse);
            this.groupBox6.Controls.Add(this.txtTreeRegionalBiomass);
            this.groupBox6.Location = new System.Drawing.Point(23, 207);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(313, 56);
            this.groupBox6.TabIndex = 37;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Tree Regional Biomass Data";
            // 
            // btnTreeRegionalBiomassBrowse
            // 
            this.btnTreeRegionalBiomassBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnTreeRegionalBiomassBrowse.Image")));
            this.btnTreeRegionalBiomassBrowse.Location = new System.Drawing.Point(272, 16);
            this.btnTreeRegionalBiomassBrowse.Name = "btnTreeRegionalBiomassBrowse";
            this.btnTreeRegionalBiomassBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnTreeRegionalBiomassBrowse.TabIndex = 1;
            this.btnTreeRegionalBiomassBrowse.Click += new System.EventHandler(this.btnTreeRegionalBiomassBrowse_Click);
            // 
            // txtTreeRegionalBiomass
            // 
            this.txtTreeRegionalBiomass.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtTreeRegionalBiomass.Location = new System.Drawing.Point(17, 19);
            this.txtTreeRegionalBiomass.Name = "txtTreeRegionalBiomass";
            this.txtTreeRegionalBiomass.Size = new System.Drawing.Size(247, 26);
            this.txtTreeRegionalBiomass.TabIndex = 0;
            // 
            // btnFIADBTxtInputFinish
            // 
            this.btnFIADBTxtInputFinish.Enabled = false;
            this.btnFIADBTxtInputFinish.Location = new System.Drawing.Point(584, 326);
            this.btnFIADBTxtInputFinish.Name = "btnFIADBTxtInputFinish";
            this.btnFIADBTxtInputFinish.Size = new System.Drawing.Size(72, 24);
            this.btnFIADBTxtInputFinish.TabIndex = 36;
            this.btnFIADBTxtInputFinish.Text = "Append";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.btnTreeBrowse);
            this.groupBox5.Controls.Add(this.txtTree);
            this.groupBox5.Location = new System.Drawing.Point(24, 143);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(312, 56);
            this.groupBox5.TabIndex = 30;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Tree Data";
            // 
            // btnTreeBrowse
            // 
            this.btnTreeBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnTreeBrowse.Image")));
            this.btnTreeBrowse.Location = new System.Drawing.Point(272, 16);
            this.btnTreeBrowse.Name = "btnTreeBrowse";
            this.btnTreeBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnTreeBrowse.TabIndex = 1;
            this.btnTreeBrowse.Click += new System.EventHandler(this.btnTreeBrowse_Click);
            // 
            // txtTree
            // 
            this.txtTree.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtTree.Location = new System.Drawing.Point(17, 19);
            this.txtTree.Name = "txtTree";
            this.txtTree.Size = new System.Drawing.Size(247, 26);
            this.txtTree.TabIndex = 0;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.btnCondBrowse);
            this.groupBox4.Controls.Add(this.txtCond);
            this.groupBox4.Location = new System.Drawing.Point(24, 80);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(312, 57);
            this.groupBox4.TabIndex = 29;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Condition Data";
            // 
            // btnCondBrowse
            // 
            this.btnCondBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnCondBrowse.Image")));
            this.btnCondBrowse.Location = new System.Drawing.Point(272, 16);
            this.btnCondBrowse.Name = "btnCondBrowse";
            this.btnCondBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnCondBrowse.TabIndex = 1;
            this.btnCondBrowse.Click += new System.EventHandler(this.btnCondBrowse_Click);
            // 
            // txtCond
            // 
            this.txtCond.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCond.Location = new System.Drawing.Point(17, 20);
            this.txtCond.Name = "txtCond";
            this.txtCond.Size = new System.Drawing.Size(247, 26);
            this.txtCond.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnPlotBrowse);
            this.groupBox2.Controls.Add(this.txtPlot);
            this.groupBox2.Location = new System.Drawing.Point(24, 17);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(312, 58);
            this.groupBox2.TabIndex = 28;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Plot Data";
            // 
            // btnPlotBrowse
            // 
            this.btnPlotBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnPlotBrowse.Image")));
            this.btnPlotBrowse.Location = new System.Drawing.Point(272, 16);
            this.btnPlotBrowse.Name = "btnPlotBrowse";
            this.btnPlotBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnPlotBrowse.TabIndex = 1;
            this.btnPlotBrowse.Click += new System.EventHandler(this.btnPlotBrowse_Click);
            // 
            // txtPlot
            // 
            this.txtPlot.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPlot.Location = new System.Drawing.Point(17, 20);
            this.txtPlot.Name = "txtPlot";
            this.txtPlot.Size = new System.Drawing.Size(247, 26);
            this.txtPlot.TabIndex = 0;
            // 
            // btnFIADBTxtInputHelp
            // 
            this.btnFIADBTxtInputHelp.Location = new System.Drawing.Point(24, 326);
            this.btnFIADBTxtInputHelp.Name = "btnFIADBTxtInputHelp";
            this.btnFIADBTxtInputHelp.Size = new System.Drawing.Size(64, 24);
            this.btnFIADBTxtInputHelp.TabIndex = 27;
            this.btnFIADBTxtInputHelp.Text = "Help";
            // 
            // btnFIADBTxtInputPrevious
            // 
            this.btnFIADBTxtInputPrevious.Enabled = false;
            this.btnFIADBTxtInputPrevious.Location = new System.Drawing.Point(424, 326);
            this.btnFIADBTxtInputPrevious.Name = "btnFIADBTxtInputPrevious";
            this.btnFIADBTxtInputPrevious.Size = new System.Drawing.Size(72, 24);
            this.btnFIADBTxtInputPrevious.TabIndex = 26;
            this.btnFIADBTxtInputPrevious.Text = "< Previous";
            this.btnFIADBTxtInputPrevious.Click += new System.EventHandler(this.btnFIADBTxtInputPrevious_Click);
            // 
            // btnFIADBTxtInputNext
            // 
            this.btnFIADBTxtInputNext.Location = new System.Drawing.Point(496, 326);
            this.btnFIADBTxtInputNext.Name = "btnFIADBTxtInputNext";
            this.btnFIADBTxtInputNext.Size = new System.Drawing.Size(72, 24);
            this.btnFIADBTxtInputNext.TabIndex = 25;
            this.btnFIADBTxtInputNext.Text = "Next >";
            this.btnFIADBTxtInputNext.Click += new System.EventHandler(this.btnFIADBTxtInputNext_Click);
            // 
            // btnFIADBTxtInputCancel
            // 
            this.btnFIADBTxtInputCancel.Location = new System.Drawing.Point(336, 326);
            this.btnFIADBTxtInputCancel.Name = "btnFIADBTxtInputCancel";
            this.btnFIADBTxtInputCancel.Size = new System.Drawing.Size(64, 24);
            this.btnFIADBTxtInputCancel.TabIndex = 24;
            this.btnFIADBTxtInputCancel.Text = "Cancel";
            this.btnFIADBTxtInputCancel.Click += new System.EventHandler(this.btnFIADBTxtInputCancel_Click);
            // 
            // grpboxInvType
            // 
            this.grpboxInvType.Controls.Add(this.btnInvTypeFinish);
            this.grpboxInvType.Controls.Add(this.btnInvTypeHelp);
            this.grpboxInvType.Controls.Add(this.btnInvTypePrevious);
            this.grpboxInvType.Controls.Add(this.btnInvTypeNext);
            this.grpboxInvType.Controls.Add(this.btnInvTypeCancel);
            this.grpboxInvType.Controls.Add(this.groupBox3);
            this.grpboxInvType.Location = new System.Drawing.Point(16, 56);
            this.grpboxInvType.Name = "grpboxInvType";
            this.grpboxInvType.Size = new System.Drawing.Size(672, 360);
            this.grpboxInvType.TabIndex = 28;
            this.grpboxInvType.TabStop = false;
            this.grpboxInvType.Text = "Inventory Type";
            // 
            // btnInvTypeFinish
            // 
            this.btnInvTypeFinish.Enabled = false;
            this.btnInvTypeFinish.Location = new System.Drawing.Point(584, 327);
            this.btnInvTypeFinish.Name = "btnInvTypeFinish";
            this.btnInvTypeFinish.Size = new System.Drawing.Size(72, 24);
            this.btnInvTypeFinish.TabIndex = 37;
            this.btnInvTypeFinish.Text = "Append";
            // 
            // btnInvTypeHelp
            // 
            this.btnInvTypeHelp.Location = new System.Drawing.Point(16, 327);
            this.btnInvTypeHelp.Name = "btnInvTypeHelp";
            this.btnInvTypeHelp.Size = new System.Drawing.Size(64, 24);
            this.btnInvTypeHelp.TabIndex = 23;
            this.btnInvTypeHelp.Text = "Help";
            // 
            // btnInvTypePrevious
            // 
            this.btnInvTypePrevious.Enabled = false;
            this.btnInvTypePrevious.Location = new System.Drawing.Point(424, 327);
            this.btnInvTypePrevious.Name = "btnInvTypePrevious";
            this.btnInvTypePrevious.Size = new System.Drawing.Size(72, 24);
            this.btnInvTypePrevious.TabIndex = 22;
            this.btnInvTypePrevious.Text = "< Previous";
            // 
            // btnInvTypeNext
            // 
            this.btnInvTypeNext.Location = new System.Drawing.Point(496, 327);
            this.btnInvTypeNext.Name = "btnInvTypeNext";
            this.btnInvTypeNext.Size = new System.Drawing.Size(72, 24);
            this.btnInvTypeNext.TabIndex = 21;
            this.btnInvTypeNext.Text = "Next >";
            this.btnInvTypeNext.Click += new System.EventHandler(this.btnInvTypeNext_Click);
            // 
            // btnInvTypeCancel
            // 
            this.btnInvTypeCancel.Location = new System.Drawing.Point(336, 327);
            this.btnInvTypeCancel.Name = "btnInvTypeCancel";
            this.btnInvTypeCancel.Size = new System.Drawing.Size(64, 24);
            this.btnInvTypeCancel.TabIndex = 20;
            this.btnInvTypeCancel.Text = "Cancel";
            this.btnInvTypeCancel.Click += new System.EventHandler(this.btnInvTypeCancel_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.grpInputDataSourceType);
            this.groupBox3.Controls.Add(this.rdoFIADB);
            this.groupBox3.Controls.Add(this.rdoIDB);
            this.groupBox3.Location = new System.Drawing.Point(108, 59);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(464, 192);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            // 
            // grpInputDataSourceType
            // 
            this.grpInputDataSourceType.Controls.Add(this.rdoText);
            this.grpInputDataSourceType.Controls.Add(this.rdoAccess);
            this.grpInputDataSourceType.Location = new System.Drawing.Point(248, 24);
            this.grpInputDataSourceType.Name = "grpInputDataSourceType";
            this.grpInputDataSourceType.Size = new System.Drawing.Size(200, 80);
            this.grpInputDataSourceType.TabIndex = 2;
            this.grpInputDataSourceType.TabStop = false;
            this.grpInputDataSourceType.Text = "Input Data Source Type";
            // 
            // rdoText
            // 
            this.rdoText.Location = new System.Drawing.Point(8, 48);
            this.rdoText.Name = "rdoText";
            this.rdoText.Size = new System.Drawing.Size(184, 24);
            this.rdoText.TabIndex = 1;
            this.rdoText.Text = "FIADB Text Files (*.csv)";
            this.rdoText.Visible = false;
            // 
            // rdoAccess
            // 
            this.rdoAccess.Checked = true;
            this.rdoAccess.Location = new System.Drawing.Point(8, 16);
            this.rdoAccess.Name = "rdoAccess";
            this.rdoAccess.Size = new System.Drawing.Size(168, 24);
            this.rdoAccess.TabIndex = 0;
            this.rdoAccess.TabStop = true;
            this.rdoAccess.Text = "Access Tables";
            // 
            // rdoFIADB
            // 
            this.rdoFIADB.Checked = true;
            this.rdoFIADB.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoFIADB.Location = new System.Drawing.Point(56, 48);
            this.rdoFIADB.Name = "rdoFIADB";
            this.rdoFIADB.Size = new System.Drawing.Size(160, 32);
            this.rdoFIADB.TabIndex = 0;
            this.rdoFIADB.TabStop = true;
            this.rdoFIADB.Text = "FIADB Data";
            this.rdoFIADB.Click += new System.EventHandler(this.rdoFIADB_Click);
            // 
            // rdoIDB
            // 
            this.rdoIDB.Enabled = false;
            this.rdoIDB.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoIDB.Location = new System.Drawing.Point(56, 128);
            this.rdoIDB.Name = "rdoIDB";
            this.rdoIDB.Size = new System.Drawing.Size(144, 32);
            this.rdoIDB.TabIndex = 1;
            this.rdoIDB.Text = "PNW IDB Data";
            this.rdoIDB.Visible = false;
            this.rdoIDB.Click += new System.EventHandler(this.rdoIDB_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(698, 24);
            this.lblTitle.TabIndex = 27;
            this.lblTitle.Text = "Plot Data Input";
            // 
            // grpboxFilterByPlot
            // 
            this.grpboxFilterByPlot.Controls.Add(this.btnFilterByPlotFinish);
            this.grpboxFilterByPlot.Controls.Add(this.btnFilterByPlotUnselect);
            this.grpboxFilterByPlot.Controls.Add(this.btnFilterByPlotSelect);
            this.grpboxFilterByPlot.Controls.Add(this.lstFilterByPlot);
            this.grpboxFilterByPlot.Controls.Add(this.btnFilterByPlotHelp);
            this.grpboxFilterByPlot.Controls.Add(this.btnFilterByPlotPrevious);
            this.grpboxFilterByPlot.Controls.Add(this.btnFilterByPlotNext);
            this.grpboxFilterByPlot.Controls.Add(this.btnFilterByPlotCancel);
            this.grpboxFilterByPlot.Location = new System.Drawing.Point(16, 1540);
            this.grpboxFilterByPlot.Name = "grpboxFilterByPlot";
            this.grpboxFilterByPlot.Size = new System.Drawing.Size(672, 360);
            this.grpboxFilterByPlot.TabIndex = 32;
            this.grpboxFilterByPlot.TabStop = false;
            this.grpboxFilterByPlot.Text = "Filter By Plot";
            this.grpboxFilterByPlot.Enter += new System.EventHandler(this.groupBox6_Enter);
            // 
            // btnFilterByPlotFinish
            // 
            this.btnFilterByPlotFinish.Location = new System.Drawing.Point(584, 325);
            this.btnFilterByPlotFinish.Name = "btnFilterByPlotFinish";
            this.btnFilterByPlotFinish.Size = new System.Drawing.Size(72, 24);
            this.btnFilterByPlotFinish.TabIndex = 33;
            this.btnFilterByPlotFinish.Text = "Append";
            this.btnFilterByPlotFinish.Click += new System.EventHandler(this.btnFilterByPlotFinish_Click);
            // 
            // btnFilterByPlotUnselect
            // 
            this.btnFilterByPlotUnselect.Location = new System.Drawing.Point(560, 180);
            this.btnFilterByPlotUnselect.Name = "btnFilterByPlotUnselect";
            this.btnFilterByPlotUnselect.Size = new System.Drawing.Size(88, 48);
            this.btnFilterByPlotUnselect.TabIndex = 32;
            this.btnFilterByPlotUnselect.Text = "Clear All";
            this.btnFilterByPlotUnselect.Click += new System.EventHandler(this.btnFilterByPlotUnselect_Click);
            // 
            // btnFilterByPlotSelect
            // 
            this.btnFilterByPlotSelect.Location = new System.Drawing.Point(560, 116);
            this.btnFilterByPlotSelect.Name = "btnFilterByPlotSelect";
            this.btnFilterByPlotSelect.Size = new System.Drawing.Size(88, 48);
            this.btnFilterByPlotSelect.TabIndex = 31;
            this.btnFilterByPlotSelect.Text = "Select All";
            this.btnFilterByPlotSelect.Click += new System.EventHandler(this.btnFilterByPlotSelect_Click);
            // 
            // lstFilterByPlot
            // 
            this.lstFilterByPlot.CheckBoxes = true;
            this.lstFilterByPlot.FullRowSelect = true;
            this.lstFilterByPlot.GridLines = true;
            this.lstFilterByPlot.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstFilterByPlot.HideSelection = false;
            this.lstFilterByPlot.Location = new System.Drawing.Point(136, 32);
            this.lstFilterByPlot.MultiSelect = false;
            this.lstFilterByPlot.Name = "lstFilterByPlot";
            this.lstFilterByPlot.Size = new System.Drawing.Size(400, 280);
            this.lstFilterByPlot.TabIndex = 30;
            this.lstFilterByPlot.UseCompatibleStateImageBehavior = false;
            this.lstFilterByPlot.View = System.Windows.Forms.View.Details;
            // 
            // btnFilterByPlotHelp
            // 
            this.btnFilterByPlotHelp.Location = new System.Drawing.Point(16, 325);
            this.btnFilterByPlotHelp.Name = "btnFilterByPlotHelp";
            this.btnFilterByPlotHelp.Size = new System.Drawing.Size(64, 24);
            this.btnFilterByPlotHelp.TabIndex = 23;
            this.btnFilterByPlotHelp.Text = "Help";
            // 
            // btnFilterByPlotPrevious
            // 
            this.btnFilterByPlotPrevious.Location = new System.Drawing.Point(424, 325);
            this.btnFilterByPlotPrevious.Name = "btnFilterByPlotPrevious";
            this.btnFilterByPlotPrevious.Size = new System.Drawing.Size(72, 24);
            this.btnFilterByPlotPrevious.TabIndex = 22;
            this.btnFilterByPlotPrevious.Text = "< Previous";
            this.btnFilterByPlotPrevious.Click += new System.EventHandler(this.btnFilterByPlotPrevious_Click);
            // 
            // btnFilterByPlotNext
            // 
            this.btnFilterByPlotNext.Enabled = false;
            this.btnFilterByPlotNext.Location = new System.Drawing.Point(496, 325);
            this.btnFilterByPlotNext.Name = "btnFilterByPlotNext";
            this.btnFilterByPlotNext.Size = new System.Drawing.Size(72, 24);
            this.btnFilterByPlotNext.TabIndex = 21;
            this.btnFilterByPlotNext.Text = "Next >";
            // 
            // btnFilterByPlotCancel
            // 
            this.btnFilterByPlotCancel.Location = new System.Drawing.Point(336, 325);
            this.btnFilterByPlotCancel.Name = "btnFilterByPlotCancel";
            this.btnFilterByPlotCancel.Size = new System.Drawing.Size(64, 24);
            this.btnFilterByPlotCancel.TabIndex = 20;
            this.btnFilterByPlotCancel.Text = "Cancel";
            this.btnFilterByPlotCancel.Click += new System.EventHandler(this.btnFilterByPlotCancel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 217);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(188, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Condition proportion percent  less than";
            // 
            // cmbCondPropPercent
            // 
            this.cmbCondPropPercent.FormattingEnabled = true;
            this.cmbCondPropPercent.Location = new System.Drawing.Point(200, 214);
            this.cmbCondPropPercent.Name = "cmbCondPropPercent";
            this.cmbCondPropPercent.Size = new System.Drawing.Size(52, 21);
            this.cmbCondPropPercent.TabIndex = 8;
            this.cmbCondPropPercent.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbCondPropPercent_KeyPress);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(270, 217);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(191, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "to change from forested to nonsampled";
            // 
            // uc_plot_input
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_plot_input";
            this.Size = new System.Drawing.Size(704, 3500);
            this.groupBox1.ResumeLayout(false);
            this.grpboxMDBFiadbInput.ResumeLayout(false);
            this.groupBox24.ResumeLayout(false);
            this.groupBox23.ResumeLayout(false);
            this.groupBox23.PerformLayout();
            this.groupBox15.ResumeLayout(false);
            this.groupBox16.ResumeLayout(false);
            this.groupBox17.ResumeLayout(false);
            this.groupBox18.ResumeLayout(false);
            this.groupBox19.ResumeLayout(false);
            this.groupBox20.ResumeLayout(false);
            this.groupBox21.ResumeLayout(false);
            this.groupBox22.ResumeLayout(false);
            this.grpboxFIADBInv.ResumeLayout(false);
            this.grpboxIDBInv.ResumeLayout(false);
            this.grpboxMDBInput.ResumeLayout(false);
            this.groupBox25.ResumeLayout(false);
            this.groupBox25.PerformLayout();
            this.groupBox8.ResumeLayout(false);
            this.groupBox8.PerformLayout();
            this.groupBox9.ResumeLayout(false);
            this.groupBox9.PerformLayout();
            this.groupBox10.ResumeLayout(false);
            this.groupBox10.PerformLayout();
            this.grpboxFilterByState.ResumeLayout(false);
            this.grpboxFilter.ResumeLayout(false);
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            this.grpboxFIADBTxtInput.ResumeLayout(false);
            this.groupBox26.ResumeLayout(false);
            this.groupBox26.PerformLayout();
            this.groupBox14.ResumeLayout(false);
            this.groupBox14.PerformLayout();
            this.groupBox13.ResumeLayout(false);
            this.groupBox13.PerformLayout();
            this.groupBox12.ResumeLayout(false);
            this.groupBox12.PerformLayout();
            this.groupBox11.ResumeLayout(false);
            this.groupBox11.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.grpboxInvType.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.grpInputDataSourceType.ResumeLayout(false);
            this.grpboxFilterByPlot.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void groupBox6_Enter(object sender, System.EventArgs e)
		{
		
		}
		private void Initialize()
		{
			this.m_DialogWd = this.Width + 10;
			this.m_DialogHt = this.groupBox1.Top + this.grpboxInvType.Top + this.grpboxInvType.Height + 100 ;

		
					
			this.grpboxFilterByState.Left = this.grpboxInvType.Left;
			this.grpboxFilterByState.Width = this.grpboxInvType.Width;
			this.grpboxFilterByState.Height = this.grpboxInvType.Height;
			this.grpboxFilterByState.Top = this.grpboxInvType.Top;
			this.btnFilterByStateHelp.Location = this.btnInvTypeHelp.Location;
			this.btnFilterByStateCancel.Location = this.btnInvTypeCancel.Location;
			this.btnFilterByStatePrevious.Location = this.btnInvTypePrevious.Location;
			this.btnFilterByStateNext.Location = this.btnInvTypeNext.Location;
			this.btnFilterByStateFinish.Location = this.btnInvTypeFinish.Location;
			this.grpboxFilterByState.Visible=false;	

			this.grpboxFIADBTxtInput.Left = this.grpboxInvType.Left;
			this.grpboxFIADBTxtInput.Width = this.grpboxInvType.Width;
			this.grpboxFIADBTxtInput.Height = this.grpboxInvType.Height;
			this.grpboxFIADBTxtInput.Top = this.grpboxInvType.Top;
			this.btnFIADBTxtInputHelp.Location = this.btnInvTypeHelp.Location;
			this.btnFIADBTxtInputCancel.Location = this.btnInvTypeCancel.Location;
			this.btnFIADBTxtInputPrevious.Location = this.btnInvTypePrevious.Location;
			this.btnFIADBTxtInputNext.Location = this.btnInvTypeNext.Location;
			this.btnFIADBTxtInputFinish.Location = this.btnInvTypeFinish.Location;
			this.grpboxFIADBTxtInput.Visible=false;	

			this.grpboxFilter.Left = this.grpboxInvType.Left;
			this.grpboxFilter.Width = this.grpboxInvType.Width;
			this.grpboxFilter.Height = this.grpboxInvType.Height;
			this.grpboxFilter.Top = this.grpboxInvType.Top;
			this.btnFilterHelp.Location = this.btnInvTypeHelp.Location;
			this.btnFilterCancel.Location = this.btnInvTypeCancel.Location;
			this.btnFilterPrevious.Location = this.btnInvTypePrevious.Location;
			this.btnFilterNext.Location = this.btnInvTypeNext.Location;
			this.btnFilterFinish.Location = this.btnInvTypeFinish.Location;
			this.grpboxFilter.Visible=false;	

			this.grpboxFilterByPlot.Left = this.grpboxInvType.Left;
			this.grpboxFilterByPlot.Width = this.grpboxInvType.Width;
			this.grpboxFilterByPlot.Height = this.grpboxInvType.Height;
			this.grpboxFilterByPlot.Top = this.grpboxInvType.Top;
			this.btnFilterByPlotHelp.Location = this.btnInvTypeHelp.Location;
			this.btnFilterByPlotCancel.Location = this.btnInvTypeCancel.Location;
			this.btnFilterByPlotPrevious.Location = this.btnInvTypePrevious.Location;
			this.btnFilterByPlotNext.Location = this.btnInvTypeNext.Location;
			this.btnFilterByPlotFinish.Location = this.btnInvTypeFinish.Location;
			this.grpboxFilterByPlot.Visible=false;


			this.grpboxMDBInput.Left = this.grpboxInvType.Left;
			this.grpboxMDBInput.Width = this.grpboxInvType.Width;
			this.grpboxMDBInput.Height = this.grpboxInvType.Height;
			this.grpboxMDBInput.Top = this.grpboxInvType.Top;
			this.btnMDBInputHelp.Location = this.btnInvTypeHelp.Location;
			this.btnMDBInputCancel.Location = this.btnInvTypeCancel.Location;
			this.btnMDBInputPrevious.Location = this.btnInvTypePrevious.Location;
			this.btnMDBInputNext.Location = this.btnInvTypeNext.Location;
			this.btnMDBInputFinish.Location = this.btnInvTypeFinish.Location;
			this.grpboxMDBInput.Visible=false;	

			this.grpboxIDBInv.Left = this.grpboxInvType.Left;
			this.grpboxIDBInv.Width = this.grpboxInvType.Width;
			this.grpboxIDBInv.Height = this.grpboxInvType.Height;
			this.grpboxIDBInv.Top = this.grpboxInvType.Top;
			this.btnIDBInvHelp.Location = this.btnInvTypeHelp.Location;
			this.btnIDBInvCancel.Location = this.btnInvTypeCancel.Location;
			this.btnIDBInvPrevious.Location = this.btnInvTypePrevious.Location;
			this.btnIDBInvNext.Location = this.btnInvTypeNext.Location;
			this.btnIDBInvAppend.Location = this.btnInvTypeFinish.Location;
			this.grpboxIDBInv.Visible=false;	

			this.grpboxFIADBInv.Left = this.grpboxInvType.Left;
			this.grpboxFIADBInv.Width = this.grpboxInvType.Width;
			this.grpboxFIADBInv.Height = this.grpboxInvType.Height;
			this.grpboxFIADBInv.Top = this.grpboxInvType.Top;
			this.btnFIADBInvHelp.Location = this.btnInvTypeHelp.Location;
			this.btnFIADBInvCancel.Location = this.btnInvTypeCancel.Location;
			this.btnFIADBInvPrevious.Location = this.btnInvTypePrevious.Location;
			this.btnFIADBInvNext.Location = this.btnInvTypeNext.Location;
			this.btnFIADBInvAppend.Location = this.btnInvTypeFinish.Location;
			this.grpboxFIADBInv.Visible=false;	


			this.grpboxMDBFiadbInput.Left = this.grpboxInvType.Left;
			this.grpboxMDBFiadbInput.Width = this.grpboxInvType.Width;
			this.grpboxMDBFiadbInput.Height = this.grpboxInvType.Height;
			this.grpboxMDBFiadbInput.Top = this.grpboxInvType.Top;
			this.btnMDBFiadbInputHelp.Location = this.btnInvTypeHelp.Location;
			this.btnMDBFiadbInputCancel.Location = this.btnInvTypeCancel.Location;
			this.btnMDBFiadbInputPrev.Location = this.btnInvTypePrevious.Location;
			this.btnMDBFiadbInputNext.Location = this.btnInvTypeNext.Location;
			this.btnMDBFiadbInputFinish.Location = this.btnInvTypeFinish.Location;
			this.grpboxMDBFiadbInput.Visible=false;	


			this.lstFilterByState.Clear();
			this.lstFilterByState.Columns.Add(" ", 100, HorizontalAlignment.Center); 
			this.lstFilterByState.Columns.Add("State", 100, HorizontalAlignment.Left);
			this.lstFilterByState.Columns.Add("County", 100, HorizontalAlignment.Left);

			//create state,count table
			
			this.m_dtStateCounty = new DataTable("statecounty");
			this.m_dtStateCounty.Columns.Add("statecd",typeof(string));
			this.m_dtStateCounty.Columns.Add("countycd",typeof(string));

			// two columns in the Primary Key.
			DataColumn[] colPk = new DataColumn[2];
			colPk[0] = this.m_dtStateCounty.Columns["statecd"];
			colPk[1] = this.m_dtStateCounty.Columns["countycd"];
			this.m_dtStateCounty.PrimaryKey = colPk;


			//create state,county,plot table
			this.m_dtPlot = new DataTable("statecountyplot");
			this.m_dtPlot.Columns.Add("statecd",typeof(string));
			this.m_dtPlot.Columns.Add("countycd",typeof(string));
			this.m_dtPlot.Columns.Add("plot",typeof(string));


			this.txtPlot.Text = @"C:\FIA_Biosum\FIADB_DATA\Oregon\OR_PLOT.CSV";
			this.txtPopEstUnit.Text = @"C:\FIA_Biosum\FIADB_DATA\Oregon\OR_POP_ESTN_UNIT.CSV";
			this.txtPopEval.Text = @"C:\FIA_Biosum\FIADB_DATA\Oregon\OR_POP_EVAL.CSV";
			this.txtPopStratum.Text = @"C:\FIA_Biosum\FIADB_DATA\Oregon\OR_POP_STRATUM.CSV";
			this.txtPpsa.Text = @"C:\FIA_Biosum\FIADB_DATA\Oregon\OR_POP_PLOT_STRATUM_ASSGN.CSV";
			this.txtTree.Text = @"C:\FIA_Biosum\FIADB_DATA\Oregon\OR_TREE.CSV";
			this.txtTreeRegionalBiomass.Text = @"C:\FIA_Biosum\FIADB_DATA\Oregon\OR_TREE_REGIONAL_BIOMASS.CSV";
			this.txtCond.Text = @"C:\FIA_Biosum\FIADB_DATA\Oregon\OR_COND.CSV";
			this.txtSiteTree.Text = @"C:\FIA_Biosum\FIADB_DATA\Oregon\OR_SITETREE.CSV";

			this.m_strCondTxtInputFile = this.txtCond.Text;
			this.m_strPlotTxtInputFile = this.txtPlot.Text;
			this.m_strPopEstUnitTxtInputFile = this.txtPopEstUnit.Text;
			this.m_strPopEvalTxtInputFile=this.txtPopEval.Text;
			this.m_strPopStratumTxtInputFile=this.txtPopStratum.Text;
			this.m_strPpsaTxtInputFile = this.txtPpsa.Text;
			this.m_strTreeTxtInputFile = this.txtTree.Text;
			this.m_strTreeRegionalBiomassTxtInputFile = this.txtTreeRegionalBiomass.Text;
			this.m_strSiteTreeTxtInputFile = this.txtSiteTree.Text;


            for (int x = 1; x <= 99; x++)
            {
                cmbCondPropPercent.Items.Add(x.ToString().Trim());
            }
            cmbCondPropPercent.Text = "25";

			


		}
        private void InitializeDatasource()
		{
			string strProjDir=frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
			
			m_oDatasource = new Datasource();
			m_oDatasource.LoadTableColumnNamesAndDataTypes=false;
			m_oDatasource.LoadTableRecordCount=false;
			m_oDatasource.m_strDataSourceMDBFile = strProjDir.Trim() + "\\db\\project.mdb";
			m_oDatasource.m_strDataSourceTableName = "datasource";
			m_oDatasource.m_strScenarioId="";
			m_oDatasource.populate_datasource_array();

			//get table names
			this.m_strPlotTable = m_oDatasource.getValidDataSourceTableName("PLOT");
			this.m_strCondTable = m_oDatasource.getValidDataSourceTableName("CONDITION");
			this.m_strTreeTable = m_oDatasource.getValidDataSourceTableName("TREE");
			this.m_strSiteTreeTable = m_oDatasource.getValidDataSourceTableName("SITE TREE");
			this.m_strTreeRegionalBiomassTable = m_oDatasource.getValidDataSourceTableName("TREE REGIONAL BIOMASS");
			this.m_strPpsaTable = m_oDatasource.getValidDataSourceTableName("POPULATION PLOT STRATUM ASSIGNMENT");
			this.m_strPopEstUnitTable = m_oDatasource.getValidDataSourceTableName("POPULATION ESTIMATION UNIT");
			this.m_strPopStratumTable = m_oDatasource.getValidDataSourceTableName("POPULATION STRATUM");
			this.m_strPopEvalTable = m_oDatasource.getValidDataSourceTableName("POPULATION EVALUATION");
            this.m_strBiosumPopStratumAdjustmentFactorsTable = m_oDatasource.getValidDataSourceTableName("BIOSUM POP STRATUM ADJUSTMENT FACTORS");
            this.m_strTreeMacroPlotBreakPointDiaTable = m_oDatasource.getValidDataSourceTableName("FIA TREE MACRO PLOT BREAKPOINT DIAMETER");
		}

		private void btnInvTypeCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}

		private void btnInvTypeNext_Click(object sender, System.EventArgs e)
		{
			if (this.m_oDatasource==null) this.InitializeDatasource();

			this.grpboxInvType.Visible=false;
			if (this.rdoFIADB.Checked==true) 
			{
				this.m_strIDBInv="";
				if (this.rdoText.Checked==true)
				{
					this.grpboxFIADBTxtInput.Visible=true;
					this.btnFIADBTxtInputPrevious.Enabled=true;
				}
				else if (this.rdoAccess.Checked==true)
				{
					this.grpboxMDBFiadbInput.Visible=true;
					this.btnMDBFiadbInputPrev.Enabled=true;
					this.btnMDBInputPrevious.Enabled=true;
				}
			}
			else
			{
				this.grpboxMDBInput.Text = "IDB Access MDB Input";
				this.grpboxMDBInput.Visible=true;
				this.btnMDBInputPrevious.Enabled=true;

			}
		}

		private void btnFIADBTxtInputPrevious_Click(object sender, System.EventArgs e)
		{
			this.grpboxFIADBTxtInput.Visible=false;
			this.grpboxInvType.Visible=true;
		}

		private void btnFIADBTxtInputCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}

		private void btnPlotBrowse_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "FIADB Plot Table Data";
			OpenFileDialog1.Filter = "Comma Delimited Text File (*.CSV;*.TXT;*.DAT) |*.csv;*.txt;*.dat";
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK) 
			{
				if (OpenFileDialog1.FileName.Trim().Length > 0) 
				{
					this.m_strPlotTxtInputFile = OpenFileDialog1.FileName.Trim();
					this.txtPlot.Text = this.m_strPlotTxtInputFile;
					//					this.m_strNewProjectFile = OpenFileDialog1.FileName.Substring(OpenFileDialog1.FileName.LastIndexOf("\\") + 1);
					//					this.m_strNewProjectDirectory = OpenFileDialog1.FileName.Substring(0,OpenFileDialog1.FileName.LastIndexOf("\\") - 3);
					//					this.OpenProjectTable(this.m_strNewProjectDirectory, this.m_strNewProjectFile);
					//					if (this.m_intError==0)
					//						this.OpenUserConfigTable(this.m_strNewProjectDirectory + "\\db", this.m_strNewProjectFile);
				}
			}
			else 
			{
				//				this.m_intError = -1;
			}
			OpenFileDialog1 = null;

		}

		private void btnCondBrowse_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "FIADB Condition Table Data";
			OpenFileDialog1.Filter = "Comma Delimited Text File (*.CSV;*.TXT;*.DAT) |*.csv;*.txt;*.dat";
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK) 
			{
				if (OpenFileDialog1.FileName.Trim().Length > 0) 
				{
					this.m_strCondTxtInputFile = OpenFileDialog1.FileName.Trim();
					this.txtCond.Text = this.m_strCondTxtInputFile;
                   
				}
			}
			else 
			{
			}
			OpenFileDialog1 = null;

		}

		private void btnTreeBrowse_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "FIADB Tree Table Data";
			OpenFileDialog1.Filter = "Comma Delimited Text File (*.CSV;*.TXT;*.DAT) |*.csv;*.txt;*.dat";
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK) 
			{
				if (OpenFileDialog1.FileName.Trim().Length > 0) 
				{
					this.m_strTreeTxtInputFile = OpenFileDialog1.FileName.Trim();
					this.txtTree.Text = this.m_strTreeTxtInputFile;
				}
			}
			else 
			{
			}
			OpenFileDialog1 = null;

		}

		private void btnFIADBTxtInputNext_Click(object sender, System.EventArgs e)
		{
			if (this.txtPlot.Text.Trim().Length == 0 ||
				this.txtCond.Text.Trim().Length == 0 || 
				this.txtTree.Text.Trim().Length == 0 ||
				this.txtTreeRegionalBiomass.Text.Trim().Length == 0 ||
				this.txtPpsa.Text.Trim().Length == 0 ||
				this.txtPopStratum.Text.Trim().Length == 0 || 
				this.txtPopEval.Text.Trim().Length == 0 || 
				this.txtPopEstUnit.Text.Trim().Length ==0 ||
				this.txtSiteTree.Text.Trim().Length==0)
			{
				MessageBox.Show("Required Data Input Missing");
			}
			else
			{
				if (this.rdoFilterByFile.Checked) this.btnFilterNext.Enabled=false;
				else this.btnFilterNext.Enabled=true;
				this.grpboxFIADBTxtInput.Visible=false;
				this.grpboxFilter.Visible=true;
				
			}
		}

		private void btnFilterPrevious_Click(object sender, System.EventArgs e)
		{
			this.grpboxFilter.Visible=false;
			if (this.rdoFIADB.Checked==true)
			{
				if (this.rdoText.Checked == true)
				{
					this.grpboxFIADBTxtInput.Visible=true;
				}
				else
				{
					this.grpboxMDBFiadbInput.Visible=true;
				}
			}
			else
			{
				this.grpboxMDBInput.Visible=true;
			}
			
		}

		private void rdoFilterByFile_Click(object sender, System.EventArgs e)
		{
			//if (rdoFilterByFile.Checked==true) 
			//{
			if (this.rdoIDB.Checked)
			{
				this.btnFilterFinish.Enabled=true;
				this.chkForested.Enabled=false;
				this.chkNonForested.Enabled=false;
				this.btnFilterNext.Enabled=false;
				this.txtFilterByFile.Enabled=true;
				this.btnFilterByFileBrowse.Enabled=true;
			}
			else
			{
				this.btnFilterFinish.Enabled=false;
				this.chkForested.Enabled=false;
				this.chkNonForested.Enabled=false;
				this.btnFilterNext.Enabled=true;
				this.txtFilterByFile.Enabled=true;
				this.btnFilterByFileBrowse.Enabled=true;
			}
			//}
		}

		private void rdoFilterByMenu_Click(object sender, System.EventArgs e)
		{
			//if (rdoFilterByFile.Checked==false && this.txtFilterByFile.Enabled==true) 
			//{
			this.btnFilterFinish.Enabled=false;
			this.chkForested.Enabled=true;
			this.chkNonForested.Enabled=true;
			this.btnFilterNext.Enabled=true;
			this.txtFilterByFile.Enabled=false;
			this.btnFilterByFileBrowse.Enabled=false;
			//}
			
		}

		private void rdoFilterNone_Click(object sender, System.EventArgs e)
		{
			//if (rdoFilterByFile.Checked==false && this.txtFilterByFile.Enabled==true) 
			//{
			if (this.rdoIDB.Checked==false)
			{
				this.btnFilterFinish.Enabled=false;
				this.chkForested.Enabled=true;
				this.chkNonForested.Enabled=true;
				this.btnFilterNext.Enabled=true;
				this.txtFilterByFile.Enabled=false;
				this.btnFilterByFileBrowse.Enabled=false;
			}
			else
			{
				this.btnFilterFinish.Enabled=false;
				this.chkForested.Enabled=true;
				this.chkNonForested.Enabled=true;
				this.btnFilterNext.Enabled=true;
				this.txtFilterByFile.Enabled=false;
				this.btnFilterByFileBrowse.Enabled=false;

			}
			//}

		}

		private void btnFilterCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}

		private void btnFilterFinish_Click(object sender, System.EventArgs e)
		{
			this.m_strStateCountyPlotSQL="";
			this.m_strStateCountySQL="";
			this.m_intError=0;
			if (this.rdoFIADB.Checked==true && this.rdoText.Checked==true)
			{
				if (this.rdoFilterNone.Checked==true) 
				{
                    LoadTxtPlotCondTreeData_Start();

				}
				else if (this.rdoFilterByFile.Checked==true)
				{
					if (System.IO.File.Exists(this.txtFilterByFile.Text.Trim()) == true)
					{
						this.m_strPlotIdList = this.CreateDelimitedStringList(this.txtFilterByFile.Text.Trim(), ",", "," , false);
						if (this.m_intError==0)
						{
                            LoadTxtPlotCondTreeData_Start();
							
						}
					}
					else
					{
						MessageBox.Show("!!" + this.txtFilterByFile.Text.Trim() + " could not be found!!","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					}
				}
			}
			else if (this.rdoFIADB.Checked==true && this.rdoAccess.Checked==true)
			{

                CalculateAdjustments_Start();
                if (m_intError == 0)
                {
                    if (this.rdoFilterNone.Checked == true)
                    {
                        LoadMDBPlotCondTreeData_Start();
                    }
                    else if (this.rdoFilterByFile.Checked == true)
                    {
                        if (System.IO.File.Exists(this.txtFilterByFile.Text.Trim()) == true)
                        {
                            this.m_strPlotIdList = this.CreateDelimitedStringList(this.txtFilterByFile.Text.Trim(), ",", ",", false);
                            if (this.m_intError == 0)
                            {
                                this.LoadMDBPlotCondTreeData_Start();
                            }
                        }
                        else
                        {
                            MessageBox.Show("!!" + this.txtFilterByFile.Text.Trim() + " could not be found!!", "Add Plot Data", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                        }
                    }
                }
			}
			else
			{
				if (this.rdoFilterNone.Checked==true) 
				{
					
                    this.LoadIDBPlotCondTreeData_Start();
				}
				else if (this.rdoFilterByFile.Checked==true)
				{
					if (System.IO.File.Exists(this.txtFilterByFile.Text.Trim()) == true)
					{
						this.m_strPlotIdList = this.CreateDelimitedStringList(this.txtFilterByFile.Text.Trim(), ",", "," ,true);
						if (this.m_intError==0)
						{
							
                            this.LoadIDBPlotCondTreeData_Start();
						}
					}
					else
					{
						MessageBox.Show("!!" + this.txtFilterByFile.Text.Trim() + " could not be found!!","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					}
				}

			}
		}
        private void LoadTxtPlotCondTreeData_Process()
		{
            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
			this.m_intError=0;
			string strFields="";
			string strCondFields="";
			string strTreeFields="";
			string strTreeRegBioFields="";
			string strSiteTreeFields="";
			int intAddedPlotRows=0;
			int intAddedCondRows=0;
			int intAddedTreeRows=0;
			int intAddedTreeRegBioRows=0;
			int intAddedSiteTreeRows=0;
			int x=0;
			string[,] strPlotArray;
			const int PLT_CN = 0;
			const int BIOSUM_PLOT_ID = 1;

			System.Data.DataTable dtTreeCN = new DataTable("TreeCN");
			dtTreeCN.Columns.Add("tre_cn",typeof(string));
			// 1 column in the Primary Key.
			DataColumn[] colTreePk = new DataColumn[1];
			colTreePk[0] =dtTreeCN.Columns["tre_cn"];
			dtTreeCN.PrimaryKey=colTreePk;
			string strCol;
			int intRecordCount=0;
					

              
			try
			{
				//instatiate the oledb data access class
				this.m_ado = new ado_data_access();
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Maximum", 6);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg2, "Text", "Overall Progress");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Visible", true);    
				    
				//create a temporary mdb file with links to all the project tables
				this.m_strTempMDBFile = m_oDatasource.CreateMDBAndTableDataSourceLinks();

				//get a connection string for the temp mdb file
				this.m_strTempMDBFileConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"","");

				//create a new connection
				this.m_connTempMDBFile = new System.Data.OleDb.OleDbConnection();

				//open the connection to the temp mdb file 
				this.m_ado.OpenConnection(this.m_strTempMDBFileConn,ref this.m_connTempMDBFile);

				//deleting previous data
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", 5);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Deleting Previous Data");
                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,"DELETE FROM " + this.m_strPlotTable + " WHERE biosum_status_cd=9");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 1);
				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,"DELETE FROM " + this.m_strCondTable + " WHERE biosum_status_cd=9");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 2);
				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,"DELETE FROM " + this.m_strTreeTable + " WHERE biosum_status_cd=9");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 3);
				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,"DELETE FROM " + this.m_strTreeRegionalBiomassTable + " WHERE biosum_status_cd=9");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 4);
				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,"DELETE FROM " + this.m_strSiteTreeTable + " WHERE biosum_status_cd=9");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 5);



                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", 3);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 1);

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				//plot table schema
				System.Data.DataTable p_dtPlotSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strPlotTable);
				strFields = "";
				for (x=0; x<=p_dtPlotSchema.Rows.Count-1;x++)
				{
					strCol = p_dtPlotSchema.Rows[x]["columnname"].ToString().Trim();
					
					if (strFields.Trim().Length == 0)
					{
						strFields = strCol;
					}
					else
					{	
						strFields += "," + strCol;
					}
				}
				System.Data.DataRow oRow = p_dtPlotSchema.NewRow();
				oRow[0] = "cycle";
				oRow[5] = "System.Int32";
				p_dtPlotSchema.Rows.Add(oRow);

				oRow = p_dtPlotSchema.NewRow();
				oRow[0] = "subcycle";
				oRow[5] = "System.Int32";
				p_dtPlotSchema.Rows.Add(oRow);

				oRow = p_dtPlotSchema.NewRow();
				oRow[0] = "plot_status_cd";
				oRow[5] = "System.Int32";
				p_dtPlotSchema.Rows.Add(oRow);
				

				string[] strPlotColumns = new string[p_dtPlotSchema.Rows.Count];
				string[] strPlotStringDataTypeYN = new string[p_dtPlotSchema.Rows.Count];

				for (x=0;x<=p_dtPlotSchema.Rows.Count-1;x++)
				{
					strPlotColumns[x] = p_dtPlotSchema.Rows[x][0].ToString().Trim();
					strPlotStringDataTypeYN[x] = m_ado.getIsTheFieldAStringDataType(p_dtPlotSchema.Rows[x][5].ToString());
				}

				//condition table schema
				System.Data.DataTable p_dtCondSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strCondTable);
				strCondFields = "";
				for (x=0; x<=p_dtCondSchema.Rows.Count-1;x++)
				{
					strCol = p_dtCondSchema.Rows[x]["columnname"].ToString().Trim();
					
					if (strCondFields.Trim().Length == 0)
					{
						strCondFields = strCol;
					}
					else
					{	
						strCondFields += "," + strCol;
					}
				}
				oRow = p_dtCondSchema.NewRow();
				oRow["columnname"] = "cond_status_cd";
				oRow["datatype"] = "System.Int32";
				p_dtCondSchema.Rows.Add(oRow);
				oRow = p_dtCondSchema.NewRow();
				oRow["columnname"] = "plt_cn";
				oRow["datatype"] = "System.String";
				oRow["columnsize"] = 34;
				oRow["AllowDbNull"] = true;
				p_dtCondSchema.Rows.Add(oRow);

				string[] strCondColumns = new string[p_dtCondSchema.Rows.Count];
				string[] strCondStringDataTypeYN = new string[p_dtCondSchema.Rows.Count];
				for (x=0;x<=p_dtCondSchema.Rows.Count-1;x++)
				{
					strCondColumns[x] = p_dtCondSchema.Rows[x][0].ToString().Trim();
					strCondStringDataTypeYN[x] = m_ado.getIsTheFieldAStringDataType(p_dtCondSchema.Rows[x][5].ToString());
				}

				//tree table schema
				System.Data.DataTable p_dtTreeSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strTreeTable);
				strTreeFields = "";
				for (x=0; x<=p_dtTreeSchema.Rows.Count-1;x++)
				{
					strCol = p_dtTreeSchema.Rows[x]["columnname"].ToString().Trim();
					
					if (strTreeFields.Trim().Length == 0)
					{
						strTreeFields = strCol;
					}
					else
					{	
						strTreeFields += "," + strCol;
					}
				}
				oRow = p_dtTreeSchema.NewRow();
				oRow["columnname"] = "plt_cn";
				oRow["datatype"] = "System.String";
				oRow["columnsize"] = 34;
				oRow["AllowDbNull"] = true;
				p_dtTreeSchema.Rows.Add(oRow);

				string[] strTreeColumns = new string[p_dtTreeSchema.Rows.Count];
				string[] strTreeStringDataTypeYN = new string[p_dtTreeSchema.Rows.Count];
				for (x=0;x<=p_dtTreeSchema.Rows.Count-1;x++)
				{
					strTreeColumns[x] = p_dtTreeSchema.Rows[x][0].ToString().Trim();
					strTreeStringDataTypeYN[x] = m_ado.getIsTheFieldAStringDataType(p_dtTreeSchema.Rows[x][5].ToString());
				}

				//tree regional biomass table
				System.Data.DataTable p_dtTreeRegBioSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strTreeRegionalBiomassTable);
				strTreeRegBioFields = "";
				for (x=0; x<=p_dtTreeRegBioSchema.Rows.Count-1;x++)
				{
					strCol = p_dtTreeRegBioSchema.Rows[x]["columnname"].ToString().Trim();
					
					if (strTreeRegBioFields.Trim().Length == 0)
					{
						strTreeRegBioFields = strCol;
					}
					else
					{	
						strTreeRegBioFields += "," + strCol;
					}
				}

				string[] strTreeRegBioColumns = new string[p_dtTreeRegBioSchema.Rows.Count];
				string[] strTreeRegBioStringDataTypeYN = new string[p_dtTreeRegBioSchema.Rows.Count];
				for (x=0;x<=p_dtTreeRegBioSchema.Rows.Count-1;x++)
				{
					strTreeRegBioColumns[x] = p_dtTreeRegBioSchema.Rows[x][0].ToString().Trim();
					strTreeRegBioStringDataTypeYN[x] = m_ado.getIsTheFieldAStringDataType(p_dtTreeRegBioSchema.Rows[x][5].ToString());
				}

				//site tree table schema
				System.Data.DataTable p_dtSiteTreeSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strSiteTreeTable);
				strSiteTreeFields = "";
				for (x=0; x<=p_dtSiteTreeSchema.Rows.Count-1;x++)
				{
					strCol = p_dtSiteTreeSchema.Rows[x]["columnname"].ToString().Trim();
					
					if (strSiteTreeFields.Trim().Length == 0)
					{
						strSiteTreeFields = strCol;
					}
					else
					{	
						strSiteTreeFields += "," + strCol;
					}
				}
				oRow = p_dtSiteTreeSchema.NewRow();
				oRow["columnname"] = "plt_cn";
				oRow["datatype"] = "System.String";
				oRow["columnsize"] = 34;
				oRow["AllowDbNull"] = true;
				p_dtSiteTreeSchema.Rows.Add(oRow);

				string[] strSiteTreeColumns = new string[p_dtSiteTreeSchema.Rows.Count];
				string[] strSiteTreeStringDataTypeYN = new string[p_dtSiteTreeSchema.Rows.Count];
				for (x=0;x<=p_dtSiteTreeSchema.Rows.Count-1;x++)
				{
					strSiteTreeColumns[x] = p_dtSiteTreeSchema.Rows[x][0].ToString().Trim();
					strSiteTreeStringDataTypeYN[x] = m_ado.getIsTheFieldAStringDataType(p_dtSiteTreeSchema.Rows[x][5].ToString());
				}



				this.m_strSQL = "SELECT biosum_cond_id, qmd_tot_cm,hwd_qmd_tot_cm," + 
					"swd_qmd_tot_cm,tpacurr,hwd_tpacurr,swd_tpacurr,ba_ft2_ac," +  
					"hwd_ba_ft2_ac,swd_ba_ft2_ac,vol_ac_grs_stem_ttl_ft3," +
					"hwd_vol_ac_grs_stem_ttl_ft3,swd_vol_ac_grs_stem_ttl_ft3," +
					"vol_ac_grs_ft3, hwd_vol_ac_grs_ft3," + 
					"swd_vol_ac_grs_ft3,volcsgrs," +
					"hwd_volcsgrs, swd_volcsgrs " + 
					"FROM " + this.m_strCondTable.Trim() + ";";


				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dtCondWorkTable = this.m_ado.getTableSchema(this.m_connTempMDBFile,this.m_strSQL);


				this.m_strSQL = "SELECT biosum_plot_id, statecd as cond_ttl " + 
					"FROM " + this.m_strPlotTable.Trim() + ";";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dtPlotWorkTable = this.m_ado.getTableSchema(this.m_connTempMDBFile,this.m_strSQL);

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 2);
			

				//close the connection to the temp mdb file
				this.m_connTempMDBFile.Close();

				/*****************************************************************
				 **create the table structure of the plot table and give it 
				 **the name of plot_input
				 *****************************************************************/
				dao_data_access p_dao = new dao_data_access();
				//plot table schema
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"plot_input",p_dtPlotSchema,true);
				if (p_dao.m_intError!=0)
				{
					this.m_intError=p_dao.m_intError;
					p_dao=null;
					p_dtPlotSchema.Dispose();
					return;
				}
				//condition table schema
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"cond_input",p_dtCondSchema,true);
				if (p_dao.m_intError!=0)
				{
					this.m_intError=p_dao.m_intError;
					p_dao=null;
					p_dtPlotSchema.Dispose();
					p_dtCondSchema.Dispose();
					return;
				}
				//tree table schema
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"tree_input",p_dtTreeSchema,true);
				if (p_dao.m_intError!=0)
				{
					this.m_intError=p_dao.m_intError;
					p_dao=null;
					p_dtPlotSchema.Dispose();
					p_dtCondSchema.Dispose();
					p_dtTreeSchema.Dispose();
					return;
				}
				//tree regional biomass table schema
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"treeRegBio_input",p_dtTreeRegBioSchema ,true);
				if (p_dao.m_intError!=0)
				{
					this.m_intError=p_dao.m_intError;
					p_dao=null;
					p_dtPlotSchema.Dispose();
					p_dtCondSchema.Dispose();
					p_dtTreeSchema.Dispose();
					p_dtTreeRegBioSchema.Dispose();
					return;
				}
				//site tree table schema
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"site_tree_input",p_dtSiteTreeSchema,true);
				if (p_dao.m_intError!=0)
				{
					this.m_intError=p_dao.m_intError;
					p_dao=null;
					p_dtPlotSchema.Dispose();
					p_dtCondSchema.Dispose();
					p_dtTreeSchema.Dispose();
					p_dtTreeRegBioSchema.Dispose();
					p_dtSiteTreeSchema.Dispose();
					return;
				}



				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"cond_column_updates_work_table",p_dtCondWorkTable,true);

				p_dtCondWorkTable.Clear();
				p_dtCondWorkTable = null;

				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"plot_column_updates_work_table",p_dtPlotWorkTable,true);

				p_dtPlotWorkTable.Clear();
				p_dtPlotWorkTable = null;


				p_dao=null;
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 3);


				//reopen the connection to the temp mdb file 
				this.m_ado.OpenConnection(this.m_strTempMDBFileConn,ref this.m_connTempMDBFile);


                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 1);

				if (this.m_intError==0)
				{

					//----------------PLOT DATA---------------//
					try
					{
						intAddedPlotRows=0;

						//see if user wanted to filter by file containing plot CN numbers
						if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.RadioButton)rdoFilterByFile,"Checked",false)
                             == true && m_strPlotIdList.Trim().Length > 0)
						{
							string strDelimiter=",";
							string[] strPlotIdArray = m_strPlotIdList.Split(strDelimiter.ToCharArray());
							this.m_ado.m_strSQL = "SELECT CN INTO input_cn FROM " + this.m_strPlotTable + " WHERE 1=2";
							this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);
							for (x=0;x<=strPlotIdArray.Length-1;x++)
							{
								if (strPlotIdArray[x] != null && strPlotIdArray[x].Trim().Length > 0)
								{
									this.m_ado.m_strSQL = "INSERT INTO input_cn (CN) VALUES ("+ strPlotIdArray[x].Trim() + ")";  
									this.m_ado.SqlNonQuery(this.m_connTempMDBFile,m_ado.m_strSQL);
								}
							}
							this.m_ado.m_strSQL = "SELECT DISTINCT a.plt_cn INTO ppsa_plt_cn_work_table " + 
								"FROM " + this.m_strPpsaTable + " a,input_cn b " + 
								"WHERE TRIM(a.plt_cn)=TRIM(b.CN) AND " + 
								"a.rscd=" + this.m_strCurrFIADBRsCd + " AND " + 
								"a.evalid=" + this.m_strCurrFIADBEvalId + " AND " + 
								"a.biosum_status_cd=9";

						}
						else
						{
				

							//copy the plot records from the ppsa table for the user selected evaluation
							this.m_ado.m_strSQL = "SELECT DISTINCT plt_cn INTO ppsa_plt_cn_work_table FROM " + this.m_strPpsaTable + " " + 
								"WHERE rscd=" + this.m_strCurrFIADBRsCd + " AND " + 
								"evalid=" + this.m_strCurrFIADBEvalId + " AND " + 
								"biosum_status_cd=9";
						}
				
				
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);

						intRecordCount = Convert.ToInt32(this.m_ado.getRecordCount(this.m_connTempMDBFile,"SELECT COUNT(*) FROM ppsa_plt_cn_work_table","ppsa_plt_cn_work_table"));

                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", intRecordCount);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Plot Records");


						//insert plot text file data into the plot_input table
						this.txtInsertFIADBDataMartDelimitedText(this.m_connTempMDBFile,
							this.txtPlot.Text.Trim(),
							"ppsa_plt_cn_work_table",
							"plt_cn",
							"plot_input",
							strPlotColumns,
							strPlotStringDataTypeYN,true);
						
						if (m_intError==0 && m_ado.m_intError==0)
						{

							//filter plot records that are part of the evaluation based on user 
							//forested,non-forested,state,county, plot selections
							this.m_ado.m_strSQL = "SELECT p.* INTO plot_input_work_table FROM plot_input p," + 
								"ppsa_plt_cn_work_table ppsa " + 
								"WHERE trim(p.cn)=trim(ppsa.plt_cn)";

							if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.RadioButton)rdoFilterByMenu,"Checked",false)
                                ==true && this.m_strStateCountySQL.Trim().Length > 0)
							{
								this.BuildFilterByStateCountyString("p.statecd","p.countycd",false);
								this.m_ado.m_strSQL = this.m_ado.m_strSQL + " AND " + this.m_strStateCountySQL;
							}
							else if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.RadioButton)rdoFilterByMenu,"Checked",false)
                                ==true && this.m_strStateCountyPlotSQL.Trim().Length > 0)
							{
								this.BuildFilterByPlotString("p.statecd","p.countycd","p.plot",false);
								this.m_ado.m_strSQL = this.m_ado.m_strSQL + " AND " + this.m_strStateCountyPlotSQL;
							}

							if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.RadioButton)rdoFilterNone,"Checked",false)
                                ==true || 
								(bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.RadioButton)rdoFilterByMenu,"Checked",false)
                                ==true)
							{

								if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.CheckBox)chkForested,"Checked",false)
                                    ==true &&
									(bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.CheckBox)chkNonForested,"Checked",false)
                                    ==false)								
								{
									this.m_ado.m_strSQL = this.m_ado.m_strSQL + " AND p.plot_status_cd=1";
								}
                                else if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.CheckBox)chkForested, "Checked", false)
                                    == false &&
                                    (bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.CheckBox)chkNonForested, "Checked", false)
                                    == true)
								{
									this.m_ado.m_strSQL = this.m_ado.m_strSQL + " AND p.plot_status_cd<>1";
								}
							}

							this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);

							if (m_intError==0 && m_ado.m_intError==0)
							{
								intRecordCount = Convert.ToInt32(this.m_ado.getRecordCount(this.m_connTempMDBFile,"SELECT COUNT(*) FROM plot_input_work_table","plot_input_work_table"));

								strPlotArray = new string[intRecordCount,2];

								this.m_ado.m_strSQL="SELECT * FROM plot_input_work_table";

								this.m_ado.SqlQueryReader(this.m_connTempMDBFile,this.m_ado.m_strSQL);
								if (m_ado.m_OleDbDataReader.HasRows)
								{
									x=0;
									while (m_ado.m_OleDbDataReader.Read())
									{
					    
										strPlotArray[x,PLT_CN]=m_ado.m_OleDbDataReader["cn"].ToString().Trim();
										strPlotArray[x,BIOSUM_PLOT_ID] = this.CreateBiosumPlotId(m_ado.m_OleDbDataReader);
										x++;
									}
								}
								m_ado.m_OleDbDataReader.Close();

								//add biosum_plot_id
								for (x=0;x<=strPlotArray.Length / 2 -1;x++)
								{
									this.m_ado.m_strSQL = "UPDATE plot_input_work_table SET biosum_plot_id = '" + 
										strPlotArray[x,BIOSUM_PLOT_ID].Trim() + "' " + 
										"WHERE TRIM(CN) = '" + strPlotArray[x,PLT_CN].Trim() + "'";
									this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);
								}

								//update misc plot columns
								this.m_ado.m_strSQL = "UPDATE plot_input_work_table " + 
									"SET gis_protected_area_yn=IIF(gis_protected_area_yn IS NULL,'N',gis_protected_area_yn)," + 
									"gis_roadless_yn=IIF(gis_roadless_yn IS NULL,'N',gis_roadless_yn)," + 
									"all_cond_not_accessible_yn=IIF(all_cond_not_accessible_yn IS NULL,'N',all_cond_not_accessible_yn)," + 
									"plot_accessible_yn=IIF(plot_accessible_yn IS NULL,'Y',plot_accessible_yn)," + 
									"biosum_status_cd=IIF(biosum_status_cd IS NULL,9,biosum_status_cd)," + 
									"gis_status_id=IIF(gis_status_id IS NULL,1,gis_status_id)";
								this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);

								//insert the plot work table records into the production plot table
								this.m_ado.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " " + 
									"(" + strFields + ") " + 
									"SELECT " + strFields + " FROM plot_input_work_table";
								this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);
								intAddedPlotRows=Convert.ToInt32(m_ado.getRecordCount(this.m_connTempMDBFile,"SELECT COUNT(*) FROM " + this.m_strPlotTable + " WHERE biosum_status_cd=9",this.m_strPlotTable));
							}
						
						}
					}
					catch (Exception err_plot)
					{
						m_intError=-1;
						MessageBox.Show(err_plot.Message);
					}
				}

				
				//-------------------CONDITION DATA------------------//
				if (this.m_intError==0 && intAddedPlotRows > 0)
				{
					try
					{
						intAddedCondRows=0;
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", intRecordCount*2);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Condition Records");

						//insert plot text file data into the plot_input table
						this.txtInsertFIADBDataMartDelimitedText(this.m_connTempMDBFile,
							this.txtCond.Text.Trim(),
							"plot_input_work_table",
							"cn",
							"cond_input",
							strCondColumns,
							strCondStringDataTypeYN,false);

						if (this.m_intError==0)
						{
							//add biosum_plot_id
							this.m_ado.m_strSQL = "UPDATE cond_input a INNER JOIN " + this.m_strPlotTable + " b " + 
								"ON TRIM(a.plt_cn)=TRIM(b.cn) " + 
								"SET a.biosum_plot_id = b.biosum_plot_id, " +
								"a.biosum_cond_id=  b.biosum_plot_id + trim(cstr(a.condid)), " +
								"a.landclcd=a.cond_status_cd," + 
								"a.cond_too_far_steep_yn=IIF(a.cond_too_far_steep_yn IS NULL,'N',a.cond_too_far_steep_yn)," + 
								"a.cond_accessible_yn=IIF(a.cond_accessible_yn IS NULL,'Y',a.cond_accessible_yn)," + 
								"a.biosum_status_cd=9";
									 

							this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);

							if (this.m_intError==0 && this.m_ado.m_intError==0)
							{

								//insert the condition work table records into the condition production table
								this.m_ado.m_strSQL = "INSERT INTO " + this.m_strCondTable + " " + 
									"(" + strCondFields + ") " + 
									"SELECT " + strCondFields + " FROM cond_input";
								this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);
								intAddedCondRows=Convert.ToInt32(m_ado.getRecordCount(this.m_connTempMDBFile,"SELECT COUNT(*) FROM " + this.m_strCondTable + " WHERE biosum_status_cd=9",this.m_strCondTable));
							}
						}
					}
					catch (Exception err_cond)
					{
						m_intError=-1;
						MessageBox.Show(err_cond.Message);

					}


				}

				//-------------------TREE DATA------------------//
				if (this.m_ado.m_intError == 0 && intAddedPlotRows > 0)
				{
					try
					{
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", (intRecordCount * 2) * 20);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 3);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Tree Table");
						intAddedTreeRows=0;
					
					
						this.txtInsertFIADBDataMartDelimitedText(this.m_connTempMDBFile,
							this.txtTree.Text.Trim(),
							"plot_input_work_table",
							"cn",
							"tree_input",
							strTreeColumns,
							strTreeStringDataTypeYN,false);

						if (this.m_intError==0 && this.m_ado.m_intError==0)
						{
							//update columns

							//add biosum_cond_id
							this.m_ado.m_strSQL = "UPDATE tree_input a INNER JOIN " + this.m_strPlotTable + " b " + 
								"ON TRIM(a.plt_cn)=TRIM(b.cn) " + 
								"SET a.biosum_cond_id=  b.biosum_plot_id + trim(cstr(a.condid)), " +
								" a.cullbf = IIF(a.cullbf IS NULL,IIF(a.cull IS NOT NULL AND a.roughcull IS NOT NULL,a.cull+a.roughcull," + 
								"IIF(a.cull IS NOT NULL,a.cull,IIF(a.roughcull IS NOT NULL,a.roughcull,0))) ,a.cullbf), " + 
								"a.biosum_status_cd=9";
									 

							this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);
							if (this.m_intError==0 && this.m_ado.m_intError==0)
							{

								//insert the tree work table records into the tree production table
								this.m_ado.m_strSQL = "INSERT INTO " + this.m_strTreeTable + " " + 
									"(" + strTreeFields + ") " + 
									"SELECT " + strTreeFields + " FROM tree_input";
								this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);
							}




							if (this.m_intError==0 && this.m_ado.m_intError==0)
								intAddedTreeRows=Convert.ToInt32(m_ado.getRecordCount(this.m_connTempMDBFile,"SELECT COUNT(*) FROM " + this.m_strTreeTable + " WHERE biosum_status_cd=9",this.m_strTreeTable));
                        
						}
					}
					catch (Exception err_tree)
					{
						m_intError=-1;
						MessageBox.Show(err_tree.Message);
					}
						                                       
						                                       
				}

				//-------------------TREE REGIONAL BIOMASS DATA------------------//
				if (this.m_intError== 0 && 
					this.m_ado.m_intError == 0 && 
					intAddedPlotRows > 0)
				{
					try
					{
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", intAddedTreeRows);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 4);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Tree Regional Biomass Table");
						intAddedTreeRegBioRows=0;

						this.txtInsertFIADBDataMartDelimitedText(this.m_connTempMDBFile,
							this.txtTreeRegionalBiomass.Text.Trim(),
							"tree_input",
							"cn",
							"treeRegBio_input",
							strTreeRegBioColumns,
							strTreeRegBioStringDataTypeYN,true);

						if (this.m_intError== 0 && 
							this.m_ado.m_intError == 0)
						{

							//update biosum_status_cd
							this.m_ado.m_strSQL = "UPDATE treeRegBio_input " + 
								"SET biosum_status_cd=9";
							this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);

							if (this.m_intError==0 && this.m_ado.m_intError==0)
							{
								//insert the tree work table records into the tree production table
								this.m_ado.m_strSQL = "INSERT INTO " + this.m_strTreeRegionalBiomassTable + " " + 
									"(" + strTreeRegBioFields + ") " + 
									"SELECT " + strTreeRegBioFields + " FROM treeRegBio_input";
								this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);
							}
						}
						if (this.m_intError== 0 && 
							this.m_ado.m_intError == 0)
							intAddedTreeRegBioRows=Convert.ToInt32(m_ado.getRecordCount(this.m_connTempMDBFile,"SELECT COUNT(*) FROM " + this.m_strTreeRegionalBiomassTable + " WHERE biosum_status_cd=9",this.m_strTreeRegionalBiomassTable));
					}
					catch (Exception err_treeregbio)
					{
						m_intError=-1;
						MessageBox.Show(err_treeregbio.Message);
					}

				}
				
				//-------------------SITE TREE DATA------------------//
				if (this.m_ado.m_intError == 0 && intAddedPlotRows > 0)
				{
					try
					{
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", (intRecordCount * 2) * 2);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 5);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Site Tree Table");
						intAddedSiteTreeRows=0;
					
					
						this.txtInsertFIADBDataMartDelimitedText(this.m_connTempMDBFile,
							this.txtSiteTree.Text.Trim(),
							"plot_input_work_table",
							"cn",
							"site_tree_input",
							strSiteTreeColumns,
							strSiteTreeStringDataTypeYN,false);

						if (this.m_intError==0 && this.m_ado.m_intError==0)
						{
							//update columns

							//add biosum_plot_id
							this.m_ado.m_strSQL = "UPDATE site_tree_input a INNER JOIN " + this.m_strPlotTable + " b " + 
								"ON TRIM(a.plt_cn)=TRIM(b.cn) " + 
								"SET a.biosum_plot_id= b.biosum_plot_id, " +
								"a.biosum_status_cd=9";
									 

							this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);
							if (this.m_intError==0 && this.m_ado.m_intError==0)
							{

								//insert the site tree work table records into the site tree production table
								this.m_ado.m_strSQL = "INSERT INTO " + this.m_strSiteTreeTable + " " + 
									"(" + strSiteTreeFields + ") " + 
									"SELECT " + strSiteTreeFields + " FROM site_tree_input";
								this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);
							}




							if (this.m_intError==0 && this.m_ado.m_intError==0)
								intAddedSiteTreeRows=Convert.ToInt32(m_ado.getRecordCount(this.m_connTempMDBFile,"SELECT COUNT(*) FROM " + this.m_strSiteTreeTable + " WHERE biosum_status_cd=9",this.m_strSiteTreeTable));
                        
						}
					}
					catch (Exception err_tree)
					{
						m_intError=-1;
						MessageBox.Show(err_tree.Message);
					}
						                                       
						                                       
				}
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 6);
				if (this.m_ado.m_intError==0 && this.m_intError == 0 && (intAddedPlotRows > 0 || intAddedCondRows > 0 || intAddedTreeRows > 0 || intAddedTreeRegBioRows > 0))
				{
					this.UpdateColumns(this.m_ado);
					if (this.m_intError==0)
					{
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Label)m_frmTherm.lblMsg, "Text", "Done");
                        frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form)m_frmTherm, "Refresh");
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1,"Maximum",false));
						this.m_strSQL = " UPDATE " + this.m_strPlotTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = " UPDATE " + this.m_strCondTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = " UPDATE " + this.m_strTreeTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = " UPDATE " + this.m_strPopEvalTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = " UPDATE " + this.m_strTreeRegionalBiomassTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = " UPDATE " + this.m_strPopStratumTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = " UPDATE " + this.m_strPpsaTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = " UPDATE " + this.m_strPopEstUnitTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = " UPDATE " + this.m_strSiteTreeTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);

                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1, "Maximum", false));
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)m_frmTherm.btnCancel, "Visible", false);
													

						MessageBox.Show("Successfully Appended \n" + 
							intAddedPlotRows.ToString().Trim() + " Plot Records \n" + 
							intAddedCondRows.ToString().Trim() + " Condition Records \n" + 
							intAddedTreeRows.ToString().Trim() + " Tree Records \n" + 
							intAddedTreeRegBioRows.ToString().Trim() + " Tree Regional Biomass Records \n" + 
							intAddedSiteTreeRows.ToString().Trim() + " Site Tree Records","Add Plot Data");

							

					}
					else
					{
						//error occurred in the updatecolumns so delete the records
						this.m_strSQL = "DELETE FROM " + this.m_strPlotTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						//delete added cond records since error occured
						this.m_strSQL = "DELETE FROM " + this.m_strCondTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						//delete added cond records since error occured
						this.m_strSQL = "DELETE FROM " + this.m_strTreeTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						//delete added cond records since error occured
						this.m_strSQL = "DELETE FROM " + this.m_strTreeRegionalBiomassTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strPopEvalTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strPopStratumTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strPpsaTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strPopEstUnitTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strSiteTreeTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						MessageBox.Show("!!Error Occured Adding Plot Records: 0 Records Added!!","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);

					}
				}
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1, "Maximum", false));
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar2, "Maximum", false));

				this.m_connTempMDBFile.Close();
				while (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
					System.Threading.Thread.Sleep(1000);
				this.m_ado.m_DataSet.Clear();
				this.m_ado.m_DataSet.Dispose();
				this.m_ado=null;
				
				//((frmDialog)this.ParentForm).m_frmMain.Visible=true;
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Visible", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);




                LoadMDBPlotCondTreeData_Finish();
			}
			
			catch (Exception error) //(System.Threading.ThreadInterruptedException e)
			{
				//MessageBox.Show(error.Message);
				//MessageBox.Show("Threading Interruption Error " + e.Message.ToString());
			}
			finally
			{
				if (this.m_connTempMDBFile != null)
				{
					if (this.m_connTempMDBFile.State == System.Data.ConnectionState.Open)
					{
						this.m_connTempMDBFile.Close();
						while (this.m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
							System.Threading.Thread.Sleep(1000);
						this.m_connTempMDBFile=null;
					}
				}
				if (m_ado != null)
				{
					if (m_ado.m_DataSet != null)
					{
						this.m_ado.m_DataSet.Clear();
						this.m_ado.m_DataSet.Dispose();
					}
					this.m_ado = null;
				}
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);
                if (this.m_frmTherm != null)
                {
                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", false);
                }


                LoadMDBPlotCondTreeData_Finish();
			}
			
		}

		private void txtInsertFIADBDataMartDelimitedText(System.Data.OleDb.OleDbConnection p_oConn,
			string p_strSourceTextFile,
			string p_strPltCnTable,
			string p_strKeyPlotCnColumn,
			string p_strDestinationTable,
			string[] p_strDestinationTableColumnsArray,
			string[] p_strDestinationTableColumnsStringDataTypeYNArray,
			bool p_bExitOnProgressBarMax)
			                                         
		{   
			//The DataSet to Return
			//DataSet result = new DataSet();
			this.m_intError=0;
			int i = 0;
			int x;
			int y;
			int z;
			int intPos;
			string strCommaDelimiter=",";
			string strValuesList="";
			string strColumnsList="";
			int intPlotCNPointer=-1;
			int intCNPointer=-1;
			int intTreeCNPointer=-1;
			bool bAddRow=false;
			string strLine="";
			int intAddCount=0;
			int intCount=0;
            string strMsg = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Label)this.m_frmTherm.lblMsg, "Text", false);

			try
			{
				//Open the file in a stream reader.
				System.IO.StreamReader s = new System.IO.StreamReader(p_strSourceTextFile);
        
				//Split the first line into the columns       
				string[] columns = s.ReadLine().Split(strCommaDelimiter.ToCharArray());



				//remove any unwanted characters from the column name
				for (x=0;x<=columns.Length-1;x++)
				{
					columns[x] = columns[x].Replace("#","");
					columns[x] = columns[x].Replace("'","");
					columns[x] = columns[x].Replace("&","");
					columns[x] = columns[x].Replace("\"","");
					if (columns[x].Trim().ToUpper()=="CN") intCNPointer=x;
					else if (columns[x].Trim().ToUpper() == "PLT_CN") intPlotCNPointer=x;
					else if (columns[x].Trim().ToUpper() == "TRE_CN") intTreeCNPointer=x;
				}

				while ((strLine = s.ReadLine()) != null)
				{
					intCount++;
					
					if (strLine.Trim().Length > 0)
					{
						bAddRow=false;
						strValuesList="";
						strColumnsList="";
						string[] items = strLine.Split(strCommaDelimiter.ToCharArray());
						//There could be occurances of the delimiter in a text string.
						//If there are then the length of the items array will be 
						//greater than the number of table columns.
						if (items.Length != columns.Length)
						{
							intPos=0;
							//load each column value in the row individually
							//reformat the row and put it in strRow
							string strRow="";
							string strDelimiter="#";
							for (i=0;i<= columns.Length - 1;i++)
							{
								string strBuildColumnString="";
								for (x=intPos;x<=strLine.Length-1;x++)
								{
									//check if we are at the end of the row
									if (x == strLine.Length - 1)
									{
										if (strBuildColumnString.Trim().Length > 0)
										{
											strRow = strRow + strBuildColumnString + strDelimiter;
										}
										else
										{
											strRow = strRow + " " + strDelimiter;
										}
									}
										
									else 
										
									{
										//check for starting double quote
										if (strLine.Substring(x,1)=="\"")
										{
											if (strBuildColumnString.Trim().Length == 0)
											{
												//check for null value
												if (strLine.Substring(x,2)=="\"\"")
												{
													strRow = strRow + " " + strDelimiter;
													intPos = x+1;
													//find the delimiter
													for (z=intPos;z<=strLine.Length-1;z++)
													{
														if (strLine.Substring(z,strCommaDelimiter.Length)==strCommaDelimiter)
														{
															intPos=z;
															break;
														}
													}
												}
												else if (strLine.Substring(x,2)=="\"" + ",")
												{
													strRow = strRow + " " + strDelimiter;
													intPos = x+1;
												}
												else
												{
											
											
													x=x+1;
													//another check for null value
													//if (r.Substring(x+1,1) != "\"")
													//{
													//MessageBox.Show(r.Substring(x+1,1));
													//check for the ending double quote
													strBuildColumnString=strLine.Substring(x,strLine.Length - x);
													z = strBuildColumnString.IndexOf("\"",2);
													if (strBuildColumnString.Trim() == "\"\"")
													{
														//no value
														strBuildColumnString=" ";
														z = strLine.IndexOf("\"",x);

													}
													else
													{
														strBuildColumnString=strBuildColumnString.Substring(0,z);
														z = strLine.IndexOf("\"",x+1);
													}
													//MessageBox.Show(strBuildColumnString + " " + strBuildColumnString.Length.ToString());
											
													//finished with the column
													strRow=strRow + strBuildColumnString + strDelimiter;
													//find where the next delimiter is
													intPos = strLine.IndexOf(strCommaDelimiter,z);
													//}
													//else
													//{
													//null value
													//	strRow = strRow + " " + strDelimiter;
													//	intPos = x+1;
													//}
												}
												break;

											}
										}
										else if (strLine.Substring(x,strCommaDelimiter.Length) == strCommaDelimiter)
										{
											if (strBuildColumnString.Trim().Length ==0)
											{
												//check for a null value
												if (strLine.Substring(x+1,strCommaDelimiter.Length)==strCommaDelimiter)
												{
													strRow = strRow + " " + strDelimiter;
													intPos=x+1;
													break;
												}
											}
											else
											{
												strRow=strRow + strBuildColumnString + strDelimiter;
												intPos=x;
												break;
											}
										}
										else
										{
									
											strBuildColumnString=strBuildColumnString + strLine.Substring(x,1);
										}
									}
									
								}

							}
							if (strRow.Trim().Length > 0) strRow=strRow.Substring(0,strRow.Length - strDelimiter.Length);
							
							//Split the row at the delimiter.
							string[] items2 = strRow.Split(strDelimiter.ToCharArray());

							//lets see if this is a valid Plot CN number
							if (p_strPltCnTable != null && p_strPltCnTable.Length > 0)
							{
								int intCnColumn=0;
								if (intPlotCNPointer >= 0)
								{
									intCnColumn = intPlotCNPointer;
								}
								else if (intTreeCNPointer >= 0)
								{
									intCnColumn = intTreeCNPointer;
								}
								else
								{
									intCnColumn= intCNPointer;
								}
								if (Convert.ToInt32(this.m_ado.getRecordCount(this.m_connTempMDBFile,"SELECT COUNT(*) FROM " + p_strPltCnTable + " WHERE trim(" + p_strKeyPlotCnColumn + ")='" + items2[intCnColumn].Trim() + "'",p_strPltCnTable)) > 0) bAddRow=true;
							}
							else
							{
								bAddRow=true;
							}
							if (bAddRow)
							{
								for (y=0;y<=p_strDestinationTableColumnsArray.Length-1;y++)
								{
									for (x=0;x<=columns[x].Length - 1;x++)
									{
										if (p_strDestinationTableColumnsArray[y].Trim().ToUpper()==
											columns[x].Trim().ToUpper())
										{
											strColumnsList = strColumnsList + columns[x] + ",";
											if (p_strDestinationTableColumnsStringDataTypeYNArray[y]=="Y")
											{
												strValuesList = strValuesList + "'" + items2[x] + "',";
											}
											else
											{
												if (items2[x].Trim().Length == 0)
													strValuesList = strValuesList + "null,";
												else strValuesList = strValuesList + items2[x] + ",";
											}
										}
									}
								}

							}

      							
						}
						else
						{
							//remove unwanted characters
							for (x=0;x<=items.Length-1;x++)
							{
								//double and single quotations
								items[x] = items[x].Replace("\"","");
								items[x] = items[x].Replace("'","");
							}
							//lets see if this is a valid Plot CN number
								
							//lets see if this is a valid Plot CN number
							if (p_strPltCnTable != null && p_strPltCnTable.Length > 0)
							{
								int intCnColumn=0;
								if (intPlotCNPointer >= 0)
								{
									intCnColumn = intPlotCNPointer;
								}
								else if (intTreeCNPointer >= 0)
								{
									intCnColumn = intTreeCNPointer;
								}
								else
								{
									intCnColumn= intCNPointer;
								}
								if (Convert.ToInt32(this.m_ado.getRecordCount(this.m_connTempMDBFile,"SELECT COUNT(*) FROM " + p_strPltCnTable + " WHERE trim(" + p_strKeyPlotCnColumn + ")='" + items[intCnColumn].Trim() + "'",p_strPltCnTable)) > 0)
									bAddRow=true;
								
							}
							else bAddRow=true;
							if (bAddRow)
							{
								for (y=0;y<=p_strDestinationTableColumnsArray.Length-1;y++)
								{
									for (x=0;x<=columns.Length - 1;x++)
									{
										if (p_strDestinationTableColumnsArray[y].Trim().ToUpper()==
											columns[x].Trim().ToUpper())
										{
											strColumnsList = strColumnsList + columns[x] + ",";
											if (p_strDestinationTableColumnsStringDataTypeYNArray[y]=="Y")
											{
												strValuesList = strValuesList + "'" + items[x] + "',";
											}
											else
											{
												if (items[x].Trim().Length == 0)
													strValuesList = strValuesList + "null,";
												else strValuesList = strValuesList + items[x] + ",";
											}
										}
									}
								}
							}
						}
						if (bAddRow)
						{
							strValuesList = strValuesList.Substring(0,strValuesList.Length - 1);
							strColumnsList = strColumnsList.Substring(0,strColumnsList.Length - 1);
							this.m_ado.m_strSQL = "INSERT INTO " + p_strDestinationTable + " " + 
								"(" + strColumnsList + ") VALUES "  + 
								"(" + strValuesList + ")";
							this.m_ado.SqlNonQuery(p_oConn,this.m_ado.m_strSQL);
							intAddCount=intAddCount+1;
							if (intAddCount < (int)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1,"Maximum",false))
                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1,"Value",(int)intAddCount);
							if (p_bExitOnProgressBarMax)
								if (intAddCount==(int)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1,"Maximum",false)) break;
							
						}
					}
                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Label)m_frmTherm.lblMsg, "Text", strMsg + " Text Line:" + intCount.ToString().Trim() + " Table Rows Added:" + intAddCount.ToString().Trim());
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Label)m_frmTherm.lblMsg, "Refresh");

				}
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", (int)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1, "Maximum", false));
				
			}
			catch (Exception caught)
			{

				this.m_intError=-1;
				MessageBox.Show("!!Error!! \n" + 
					"Module - ado_data_access:ConvertDelimitedTextToDataTable  \n" + 
					"Err Msg - " + caught.Message.ToString().Trim(),
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}
			if (m_intError==0) 
			{
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Label)m_frmTherm.lblMsg, "Text", strMsg + "End Of File. Total Records Added To Table:" + intAddCount.ToString());
                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form)m_frmTherm, "Refresh");
			}
		
		}


		private void txtFileInputPopEval()
		{
			//variables used to match up the biosum table columns with the 
			//fiadb txtfile table column names
			int intFIADBColumn=0;
			int intBiosumColumn=0;

			string strValue="";
			string strFields="";
			string strValues="";
			int intAddedRows=0;
			bool bAddRow=true;
			int x=0;
			System.Data.DataTable p_dtPopEvalChanges;
			System.Data.DataRow[] p_PopEvalNewRows;
			string[,] strPopEvalList;
			const int POPEVAL_CN = 0;
			string strPlotEvalCn="";
			int[] intCol;
			
              
			try
			{
				//instatiate the oledb data access class
				this.m_ado = new ado_data_access();
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Adding Pop Eval Records");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", 6);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);


                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Visible", true);

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Creating Datasource Links");
                    
				//get all the project datasources
				FIA_Biosum_Manager.Datasource p_datasource = new Datasource(((frmDialog)this.ParentForm).m_frmMain.frmProject.uc_project1.txtRootDirectory.Text.ToString().Trim());

				//create a temporary mdb file with links to all the project tables
				this.m_strTempMDBFile = p_datasource.CreateMDBAndTableDataSourceLinks();

				//get a connection string for the temp mdb file
				this.m_strTempMDBFileConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"","");

				//get the plot, cond, and tree table names
				this.m_strPopEvalTable = p_datasource.getValidDataSourceTableName("POPULATION EVALUATION");

				//remove any entries that have a biosum_status_code of 9
				this.m_ado.SqlNonQuery(this.m_strTempMDBFileConn,"DELETE FROM " + this.m_strPopEvalTable + "  WHERE biosum_status_cd = 9;");

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 1);
				   					
				//create a new connection
				this.m_connTempMDBFile = new System.Data.OleDb.OleDbConnection();

				//open the connection to the temp mdb file 
				this.m_ado.OpenConnection(this.m_strTempMDBFileConn,ref this.m_connTempMDBFile);

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dtPopEvalSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strPopEvalTable);

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 2);
                			

				//close the connection to the temp mdb file
				this.m_connTempMDBFile.Close();

				/*****************************************************************
				 **create the table structure of the pop eval table and give it 
				 **the name of popstratum_input
				 *****************************************************************/
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"popstratum_input",p_dtPopEvalSchema,true);
				if (p_dao.m_intError!=0)
				{
					this.m_intError=p_dao.m_intError;
					p_dao=null;
					p_dtPopEvalSchema.Dispose();
					return;
				}

				p_dao=null;

				//reopen the connection to the temp mdb file 
				this.m_ado.OpenConnection(this.m_strTempMDBFileConn,ref this.m_connTempMDBFile);
                   
				//initialize the transaction object with the temporary connection
				System.Data.OleDb.OleDbTransaction p_trans = this.m_connTempMDBFile.BeginTransaction();

				FIA_Biosum_Manager.env p_env = new env();
				System.Data.DataRow p_row;
				this.m_ado.m_DataSet = new DataSet("FIADB");
				this.m_ado.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 1);
				

				//----------------pop eval DATA---------------//
				//load the text file into an adodot net datasource table
				this.m_ado.ConvertDelimitedTextToDataTable(this.m_ado.m_DataSet,this.txtPopEval.Text.Trim(),"fiadb_popeval",",");

				if (this.m_ado.m_intError==0)
				{
					try
					{
						strPopEvalList = new string[this.m_ado.m_DataSet.Tables["fiadb_popeval"].Rows.Count,1];
						for (x=0;x<=this.m_ado.m_DataSet.Tables["fiadb_popeval"].Rows.Count-1;x++)
						{
							strPopEvalList[x,POPEVAL_CN]="";

						}

						if (this.m_ado.m_DataSet.Tables["fiadb_popeval"].Rows.Count > 0)
						{
								
								 
							this.m_ado.AddSQLQueryToDataSet(this.m_connTempMDBFile,ref this.m_ado.m_OleDbDataAdapter,ref this.m_ado.m_DataSet,ref p_trans, "select * from " + this.m_strPopEvalTable,this.m_strPopEvalTable);
								
							strFields = "";
							strValues = "";
							//Build the pop eval insert sql
							for (x=0; x<=p_dtPopEvalSchema.Rows.Count-1;x++)
							{
								if (strFields.Trim().Length == 0)
								{
									strFields = "(";
								}
								else
								{	
									strFields = strFields + "," ;
								}
								strFields = strFields + p_dtPopEvalSchema.Rows[x]["columnname"].ToString().Trim();
								if (strValues.Trim().Length == 0)
								{
									strValues = "(";
								}
								else
								{	
									strValues = strValues + ",";
								}
								strValues = strValues + "?";

							}
							strFields = strFields + ")";
							strValues = strValues + ");";
							//create an insert command 
							this.m_ado.m_OleDbDataAdapter.InsertCommand = this.m_connTempMDBFile.CreateCommand();
							//bind the transaction object to the insert command
							this.m_ado.m_OleDbDataAdapter.InsertCommand.Transaction = p_trans;
							this.m_ado.m_OleDbDataAdapter.InsertCommand.CommandText = 
								"INSERT INTO " + this.m_strPopEvalTable + " "  + strFields + " VALUES " + strValues;
							//define field datatypes for the data adapter
							for (x=0; x<=p_dtPopEvalSchema.Rows.Count-1;x++)
							{
								strFields=p_dtPopEvalSchema.Rows[x]["columnname"].ToString().Trim();
								switch (p_dtPopEvalSchema.Rows[x]["datatype"].ToString().Trim())
								{
									case "System.String" :
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.VarWChar,
											0,
											strFields);
										break;
									case "System.Double":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Double,
											0,
											strFields);
										break;
									case "System.Boolean":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Boolean,
											0,
											strFields);
										break;
									case "System.DateTime":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.DBTimeStamp,
											0,
											strFields);
										break;
									case "System.Decimal":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Decimal,
											0,
											strFields);
										break;
									case "System.Int16":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.SmallInt,
											0,
											strFields);
										break;
									case "System.Int32":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Integer,
											0,
											strFields);
										break;
									case "System.Int64":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.BigInt,
											0,
											strFields);
										break;
									case "System.SByte":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.SmallInt,
											0,
											strFields);
										break;
									case "System.Byte":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.TinyInt,
											0,
											strFields);
										break;
									case "System.Single":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Single,
											0,
											strFields);
										break;
									default:
										MessageBox.Show(this.m_ado.m_DataSet.Tables[this.m_strPopEvalTable].Columns[x].DataType.FullName.ToString().Trim());
										break;
								}
									
							}
								
							intCol = new int[this.m_ado.m_DataSet.Tables[this.m_strPopEvalTable].Columns.Count];

							//match up the biosum columns with the fiadb columns
							for (intBiosumColumn = 0; intBiosumColumn <= this.m_ado.m_DataSet.Tables[this.m_strPopEvalTable].Columns.Count-1;intBiosumColumn++)
							{
								intCol[intBiosumColumn]=-1;
								for (intFIADBColumn = 0; intFIADBColumn <= this.m_ado.m_DataSet.Tables["fiadb_popeval"].Columns.Count-1;intFIADBColumn++)
								{
									if (this.m_ado.m_DataSet.Tables["fiadb_popeval"].Columns[intFIADBColumn].ColumnName.Trim().ToUpper() == 
										this.m_ado.m_DataSet.Tables[this.m_strPopEvalTable].Columns[intBiosumColumn].ColumnName.Trim().ToUpper())
									{
										intCol[intBiosumColumn] = intFIADBColumn;
									}

								}
							}
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Pop Eval Table: Compiling New Pop Eval Rows");
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

							
							//load up each row in the FIADB plot input table
							for (x = 0; x<=this.m_ado.m_DataSet.Tables["fiadb_popeval"].Rows.Count-1;x++)
							{
								bAddRow=true;
								//make sure the row is not null values
								if (this.m_ado.m_DataSet.Tables["fiadb_popeval"].Rows[x][0] != System.DBNull.Value &&
									this.m_ado.m_DataSet.Tables["fiadb_popeval"].Rows[x][0].ToString().Trim().Length > 0)
								{
									
									p_row = this.m_ado.m_DataSet.Tables[this.m_strPopEvalTable].NewRow();
									for (intBiosumColumn = 0; intBiosumColumn <= intCol.Length-1; intBiosumColumn++)
									{
										if (intCol[intBiosumColumn] != -1)
										{
											strValue = this.m_ado.m_DataSet.Tables["fiadb_popeval"].Rows[x][intCol[intBiosumColumn]].ToString().Trim();
											if (strValue.Trim().Length > 0)
											{
												switch (this.m_ado.m_DataSet.Tables[this.m_strPopEvalTable].Columns[intBiosumColumn].DataType.FullName.ToString().Trim())
												{
													case "System.String" :
														p_row[intBiosumColumn] = strValue.Replace("\"","");
														break;
													case "System.Double":
														p_row[intBiosumColumn] = Convert.ToDouble(strValue);
														break;
													case "System.Boolean":
														p_row[intBiosumColumn] = Convert.ToBoolean(strValue);
														break;
													case "System.DateTime":
														p_row[intBiosumColumn] = Convert.ToDateTime(strValue);
														break;
													case "System.Decimal":
														p_row[intBiosumColumn] = Convert.ToDecimal(strValue);
														break;
													case "System.Int16":
														p_row[intBiosumColumn] = Convert.ToInt16(strValue);
														break;
													case "System.Int32":
														p_row[intBiosumColumn] = Convert.ToInt32(strValue);
														break;
													case "System.Int64":
														p_row[intBiosumColumn] = Convert.ToInt64(strValue);
														break;
													case "System.SByte":
														p_row[intBiosumColumn] = Convert.ToSByte(strValue);

														break;
													case "System.Single":
														p_row[intBiosumColumn] = Convert.ToSingle(strValue);
														break;

													default:
														MessageBox.Show(this.m_ado.m_DataSet.Tables[this.m_strPopEvalTable].Columns[intBiosumColumn].DataType.FullName.ToString().Trim());
														break;
												}
											}
										}
										System.Windows.Forms.Application.DoEvents();
										if (this.m_frmTherm.AbortProcess == true) break;
									}
									if (this.m_frmTherm.AbortProcess == false)
									{
											
										//get the plt_cn value
										strPlotEvalCn = this.m_ado.m_DataSet.Tables["fiadb_popeval"].Rows[x]["cn"].ToString().Trim();
										strPlotEvalCn = strPlotEvalCn.Replace("\"","");
										//if (strPlotEvalCn.Trim() == "17546689010497")
										//	MessageBox.Show("here it is");

										//see if user has a list to filter the plots to process
										if (this.m_strPlotIdList.Trim().Length > 0 )
										{
											if (this.m_strPlotIdList.IndexOf("'" + strPlotEvalCn.Trim() + "'") < 0)
											{
												bAddRow=false;
											}
										}
										if (bAddRow==true)
										{
											if (p_row["biosum_status_cd"] == System.DBNull.Value) 
												p_row["biosum_status_cd"] = 9;		
											strPopEvalList[intAddedRows,POPEVAL_CN]=strPlotEvalCn.Trim();
											this.m_ado.m_DataSet.Tables[this.m_strPopEvalTable].Rows.Add(p_row);
											intAddedRows++;
											
										}
									}
								}
							
							}
						}
						else
						{
							MessageBox.Show("!!No Pop Eval Records In The FIADB Pop Eval Input Table!!");
						}
					}
					catch (Exception caught)
					{
						this.m_intError=-1;
						MessageBox.Show(caught.Message);
					}
					//keep processing the records if there was no error and there were plot records added
					if (this.m_intError==0 && intAddedRows > 0 && this.m_frmTherm.AbortProcess == false)
					{
						try
						{
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Pop Eval Table: Finishing New Pop Eval Rows...Stand By");
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
                            p_dtPopEvalChanges = this.m_ado.m_DataSet.Tables[this.m_strPopEvalTable].GetChanges();
								
							p_PopEvalNewRows = p_dtPopEvalChanges.Select(null,null, DataViewRowState.Added);
							if (p_dtPopEvalChanges.HasErrors)
							{
								this.m_ado.m_DataSet.Tables[this.m_strPopEvalTable].RejectChanges();
								this.m_intError=-1;
							}
							else
							{
								this.m_ado.m_OleDbDataAdapter.Update(p_PopEvalNewRows);
								this.m_ado.m_DataSet.Tables[this.m_strPopEvalTable].AcceptChanges();
							}
						}
						catch (Exception caught)
						{
							this.m_intError=-1;
							this.m_ado.m_DataSet.Tables[this.m_strPopEvalTable].RejectChanges();
							//rollback the transaction to the original records 
							p_trans.Rollback();
							MessageBox.Show(caught.Message);
						}
					}
				}
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 4);
								
				if (this.m_frmTherm.AbortProcess==false && this.m_intError == 0 && (intAddedRows > 0))
				{
					try
					{
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 1);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Committing The Data...Stand by");
                        frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
						p_trans.Commit(); 

					}
					catch //(Exception caught)
					{
						p_trans.Rollback();
						this.m_intError=-1;
					}

				}
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 5);
				if (this.m_frmTherm.AbortProcess == true &&  (intAddedRows > 0))
				{
					p_trans.Rollback();
				
				}
				else
				{
					if (this.m_frmTherm.AbortProcess==false && this.m_ado.m_intError==0 && this.m_intError == 0 && (intAddedRows > 0))
					{
						p_trans = null;

						if (this.m_intError==0)
						{
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Done");
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", (int)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1, "Maximum", false));
						}
						else
						{
							//error occurred in the updatecolumns so delete the records
							this.m_strSQL = "DELETE FROM " + this.m_strPopEvalTable + " WHERE biosum_status_cd = 9;";
							this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
							MessageBox.Show("!!Error Occured Adding Pop Eval Records: 0 Records Added!!","Add Pop Eval Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
						}
					}

				}
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", (int)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1, "Maximum", false));
                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form)m_frmTherm, "Close");
				this.m_frmTherm = null;
				p_dtPopEvalChanges = null;
				p_PopEvalNewRows = null;
				this.m_ado.m_DataSet.Clear();
				this.m_ado.m_DataSet.Dispose();
				this.m_ado = null;
				((frmDialog)this.ParentForm).Enabled=true;

			}
			
			catch //(System.Threading.ThreadInterruptedException e)
			{
				//MessageBox.Show("Threading Interruption Error " + e.Message.ToString());
			}
		


		}
		private void txtFileInputPopStratum()
		{
			//variables used to match up the biosum table columns with the 
			//fiadb txtfile table column names
			int intFIADBColumn=0;
			int intBiosumColumn=0;

			string strValue="";
			string strFields="";
			string strValues="";
			int intAddedRows=0;
			bool bAddRow=true;
			int x=0;
			System.Data.DataTable p_dtChanges;
			System.Data.DataRow[] p_dtNewRows;
			string[,] strList;
			const int CN = 0;
			string strCn="";
			int[] intCol;
			

			try
			{
				//instatiate the oledb data access class
				this.m_ado = new ado_data_access();
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", 6);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Text", "Adding Pop Stratum Records");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Creating Datasource Links");
				    
				//get all the project datasources
				FIA_Biosum_Manager.Datasource p_datasource = new Datasource(((frmDialog)this.ParentForm).m_frmMain.frmProject.uc_project1.txtRootDirectory.Text.ToString().Trim());

				//create a temporary mdb file with links to all the project tables
				this.m_strTempMDBFile = p_datasource.CreateMDBAndTableDataSourceLinks();

				//get a connection string for the temp mdb file
				this.m_strTempMDBFileConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"","");

				//get the plot, cond, and tree table names
				this.m_strPopStratumTable = p_datasource.getValidDataSourceTableName("POPULATION STRATUM");

				//remove any entries that have a biosum_status_code of 9
				this.m_ado.SqlNonQuery(this.m_strTempMDBFileConn,"DELETE FROM " + this.m_strPopStratumTable + "  WHERE biosum_status_cd = 9;");

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 1);
   					
				//create a new connection
				this.m_connTempMDBFile = new System.Data.OleDb.OleDbConnection();

				//open the connection to the temp mdb file 
				this.m_ado.OpenConnection(this.m_strTempMDBFileConn,ref this.m_connTempMDBFile);

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dtPopStratumSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strPopStratumTable);

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 2);
			

				//close the connection to the temp mdb file
				this.m_connTempMDBFile.Close();

				/*****************************************************************
				 **create the table structure of the pop eval table and give it 
				 **the name of popstratum_input
				 *****************************************************************/
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"popstratum_input",p_dtPopStratumSchema,true);
				if (p_dao.m_intError!=0)
				{
					this.m_intError=p_dao.m_intError;
					p_dao=null;
					p_dtPopStratumSchema.Dispose();
					return;
				}

				p_dao=null;

				//reopen the connection to the temp mdb file 
				this.m_ado.OpenConnection(this.m_strTempMDBFileConn,ref this.m_connTempMDBFile);
                   
				//initialize the transaction object with the temporary connection
				System.Data.OleDb.OleDbTransaction p_trans = this.m_connTempMDBFile.BeginTransaction();

				FIA_Biosum_Manager.env p_env = new env();
				System.Data.DataRow p_row;
				this.m_ado.m_DataSet = new DataSet("FIADB");
				this.m_ado.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 1);

				//----------------pop stratum DATA---------------//
				//load the text file into an adodot net datasource table
				this.m_ado.ConvertDelimitedTextToDataTable(this.m_ado.m_DataSet,this.txtPopStratum.Text.Trim(),"fiadb_popstratum",",");

				if (this.m_ado.m_intError==0)
				{
					try
					{
						strList = new string[this.m_ado.m_DataSet.Tables["fiadb_popstratum"].Rows.Count,1];
						for (x=0;x<=this.m_ado.m_DataSet.Tables["fiadb_popstratum"].Rows.Count-1;x++)
						{
							strList[x,CN]="";

						}

						if (this.m_ado.m_DataSet.Tables["fiadb_popstratum"].Rows.Count > 0)
						{
								
								 
							this.m_ado.AddSQLQueryToDataSet(this.m_connTempMDBFile,ref this.m_ado.m_OleDbDataAdapter,ref this.m_ado.m_DataSet,ref p_trans, "select * from " + this.m_strPopStratumTable,this.m_strPopStratumTable);
								
							strFields = "";
							strValues = "";
							//Build the pop eval insert sql
							for (x=0; x<=p_dtPopStratumSchema.Rows.Count-1;x++)
							{
								if (strFields.Trim().Length == 0)
								{
									strFields = "(";
								}
								else
								{	
									strFields = strFields + "," ;
								}
								strFields = strFields + p_dtPopStratumSchema.Rows[x]["columnname"].ToString().Trim();
								if (strValues.Trim().Length == 0)
								{
									strValues = "(";
								}
								else
								{	
									strValues = strValues + ",";
								}
								strValues = strValues + "?";

							}
							strFields = strFields + ")";
							strValues = strValues + ");";
							//create an insert command 
							this.m_ado.m_OleDbDataAdapter.InsertCommand = this.m_connTempMDBFile.CreateCommand();
							//bind the transaction object to the insert command
							this.m_ado.m_OleDbDataAdapter.InsertCommand.Transaction = p_trans;
							this.m_ado.m_OleDbDataAdapter.InsertCommand.CommandText = 
								"INSERT INTO " + this.m_strPopStratumTable + " "  + strFields + " VALUES " + strValues;
							//define field datatypes for the data adapter
							for (x=0; x<=p_dtPopStratumSchema.Rows.Count-1;x++)
							{
								strFields=p_dtPopStratumSchema.Rows[x]["columnname"].ToString().Trim();
								switch (p_dtPopStratumSchema.Rows[x]["datatype"].ToString().Trim())
								{
									case "System.String" :
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.VarWChar,
											0,
											strFields);
										break;
									case "System.Double":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Double,
											0,
											strFields);
										break;
									case "System.Boolean":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Boolean,
											0,
											strFields);
										break;
									case "System.DateTime":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.DBTimeStamp,
											0,
											strFields);
										break;
									case "System.Decimal":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Decimal,
											0,
											strFields);
										break;
									case "System.Int16":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.SmallInt,
											0,
											strFields);
										break;
									case "System.Int32":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Integer,
											0,
											strFields);
										break;
									case "System.Int64":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.BigInt,
											0,
											strFields);
										break;
									case "System.SByte":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.SmallInt,
											0,
											strFields);
										break;
									case "System.Byte":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.TinyInt,
											0,
											strFields);
										break;
									case "System.Single":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Single,
											0,
											strFields);
										break;
									default:
										MessageBox.Show(this.m_ado.m_DataSet.Tables[this.m_strPopStratumTable].Columns[x].DataType.FullName.ToString().Trim());
										break;
								}
									
							}
								
							intCol = new int[this.m_ado.m_DataSet.Tables[this.m_strPopStratumTable].Columns.Count];

							//match up the biosum columns with the fiadb columns
							for (intBiosumColumn = 0; intBiosumColumn <= this.m_ado.m_DataSet.Tables[this.m_strPopStratumTable].Columns.Count-1;intBiosumColumn++)
							{
								intCol[intBiosumColumn]=-1;
								for (intFIADBColumn = 0; intFIADBColumn <= this.m_ado.m_DataSet.Tables["fiadb_popstratum"].Columns.Count-1;intFIADBColumn++)
								{
									if (this.m_ado.m_DataSet.Tables["fiadb_popstratum"].Columns[intFIADBColumn].ColumnName.Trim().ToUpper() == 
										this.m_ado.m_DataSet.Tables[this.m_strPopStratumTable].Columns[intBiosumColumn].ColumnName.Trim().ToUpper())
									{
										intCol[intBiosumColumn] = intFIADBColumn;
									}

								}
							}
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Pop Stratum Table: Compiling New Pop Statum Rows");
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Visible", true);
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
							//load up each row in the FIADB plot input table
							for (x = 0; x<=this.m_ado.m_DataSet.Tables["fiadb_popstratum"].Rows.Count-1;x++)
							{
								bAddRow=true;
								//make sure the row is not null values
								if (this.m_ado.m_DataSet.Tables["fiadb_popstratum"].Rows[x][0] != System.DBNull.Value &&
									this.m_ado.m_DataSet.Tables["fiadb_popstratum"].Rows[x][0].ToString().Trim().Length > 0)
								{
									
									p_row = this.m_ado.m_DataSet.Tables[this.m_strPopStratumTable].NewRow();
									for (intBiosumColumn = 0; intBiosumColumn <= intCol.Length-1; intBiosumColumn++)
									{
										if (intCol[intBiosumColumn] != -1)
										{
											strValue = this.m_ado.m_DataSet.Tables["fiadb_popstratum"].Rows[x][intCol[intBiosumColumn]].ToString().Trim();
											if (strValue.Trim().Length > 0)
											{
												switch (this.m_ado.m_DataSet.Tables[this.m_strPopStratumTable].Columns[intBiosumColumn].DataType.FullName.ToString().Trim())
												{
													case "System.String" :
														p_row[intBiosumColumn] = strValue.Replace("\"","");
														break;
													case "System.Double":
														p_row[intBiosumColumn] = Convert.ToDouble(strValue);
														break;
													case "System.Boolean":
														p_row[intBiosumColumn] = Convert.ToBoolean(strValue);
														break;
													case "System.DateTime":
														p_row[intBiosumColumn] = Convert.ToDateTime(strValue);
														break;
													case "System.Decimal":
														p_row[intBiosumColumn] = Convert.ToDecimal(strValue);
														break;
													case "System.Int16":
														p_row[intBiosumColumn] = Convert.ToInt16(strValue);
														break;
													case "System.Int32":
														p_row[intBiosumColumn] = Convert.ToInt32(strValue);
														break;
													case "System.Int64":
														p_row[intBiosumColumn] = Convert.ToInt64(strValue);
														break;
													case "System.SByte":
														p_row[intBiosumColumn] = Convert.ToSByte(strValue);

														break;
													case "System.Single":
														p_row[intBiosumColumn] = Convert.ToSingle(strValue);
														break;

													default:
														MessageBox.Show(this.m_ado.m_DataSet.Tables[this.m_strPopStratumTable].Columns[intBiosumColumn].DataType.FullName.ToString().Trim());
														break;
												}
											}
										}
										System.Windows.Forms.Application.DoEvents();
										if (this.m_frmTherm.AbortProcess == true) break;
									}
									if (this.m_frmTherm.AbortProcess == false)
									{
											
										//get the plt_cn value
										strCn = this.m_ado.m_DataSet.Tables["fiadb_popstratum"].Rows[x]["cn"].ToString().Trim();
										strCn = strCn.Replace("\"","");

										//see if user has a list to filter the plots to process
										if (this.m_strPlotIdList.Trim().Length > 0 )
										{
											if (this.m_strPlotIdList.IndexOf("'" + strCn.Trim() + "'") < 0)
											{
												bAddRow=false;
											}
										}
										
										//check rscd and evalid
										if (m_ado.m_DataSet.Tables["fiadb_popstratum"].Rows[x]["rscd"] != System.DBNull.Value &&
											m_ado.m_DataSet.Tables["fiadb_popstratum"].Rows[x]["evalid"] != System.DBNull.Value)
										{
											if (Convert.ToString(m_ado.m_DataSet.Tables["fiadb_popstratum"].Rows[x]["rscd"]).Trim() != this.m_strCurrFIADBRsCd.Trim() ||
												Convert.ToString(m_ado.m_DataSet.Tables["fiadb_popstratum"].Rows[x]["evalid"]).Trim() != this.m_strCurrFIADBEvalId.Trim())
											{
												bAddRow=false;
											}


										}
										if (bAddRow==true)
										{
											if (p_row["biosum_status_cd"] == System.DBNull.Value) 
												p_row["biosum_status_cd"] = 9;		
											strList[intAddedRows,CN]=strCn.Trim();
											this.m_ado.m_DataSet.Tables[this.m_strPopStratumTable].Rows.Add(p_row);
											intAddedRows++;
											
										}
									}
								}
							
							}
						}
						else
						{
							MessageBox.Show("!!No Pop Eval Records In The FIADB Pop Eval Input Table!!");
						}
					}
					catch (Exception caught)
					{
						this.m_intError=-1;
						MessageBox.Show(caught.Message);
					}
					//keep processing the records if there was no error and there were plot records added
					if (this.m_intError==0 && intAddedRows > 0 && this.m_frmTherm.AbortProcess == false)
					{
						try
						{
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Pop Eval Table: Finishing New Pop Eval Rows...Stand By");
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
                            p_dtChanges = this.m_ado.m_DataSet.Tables[this.m_strPopStratumTable].GetChanges();
								
							p_dtNewRows = p_dtChanges.Select(null,null, DataViewRowState.Added);
							if (p_dtChanges.HasErrors)
							{
								this.m_ado.m_DataSet.Tables[this.m_strPopStratumTable].RejectChanges();
								this.m_intError=-1;
							}
							else
							{
								this.m_ado.m_OleDbDataAdapter.Update(p_dtNewRows);
								this.m_ado.m_DataSet.Tables[this.m_strPopStratumTable].AcceptChanges();
							}
						}
						catch (Exception caught)
						{
							this.m_intError=-1;
							this.m_ado.m_DataSet.Tables[this.m_strPopStratumTable].RejectChanges();
							//rollback the transaction to the original records 
							p_trans.Rollback();
							MessageBox.Show(caught.Message);
						}
					}
				}
				
				this.m_frmTherm.progressBar1.Value=4;
				
				if (this.m_frmTherm.AbortProcess==false && this.m_intError == 0 && (intAddedRows > 0))
				{
					try
					{
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Committing The Data...Stand by");
                        frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
						p_trans.Commit(); 

					}
					catch //(Exception caught)
					{
						p_trans.Rollback();
						this.m_intError=-1;
					}

				}
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 5);
				if (this.m_frmTherm.AbortProcess == true &&  (intAddedRows > 0))
				{
					p_trans.Rollback();
				
				}
				else
				{
					if (this.m_frmTherm.AbortProcess==false && this.m_ado.m_intError==0 && this.m_intError == 0 && (intAddedRows > 0))
					{
						p_trans = null;

						if (this.m_intError==0)
						{
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Done");
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", (int)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1, "Maximum", false));
						}
						else
						{
							//error occurred in the updatecolumns so delete the records
							this.m_strSQL = "DELETE FROM " + this.m_strPopStratumTable + " WHERE biosum_status_cd = 9;";
							this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
							MessageBox.Show("!!Error Occured Adding Pop Eval Records: 0 Records Added!!","Add Pop Eval Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
						}
					}

				}
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", (int)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1, "Maximum", false));
                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form)m_frmTherm, "Close");
				this.m_frmTherm = null;
				p_dtChanges = null;
				p_dtNewRows = null;
				this.m_ado.m_DataSet.Clear();
				this.m_ado.m_DataSet.Dispose();
				this.m_ado = null;
				((frmDialog)this.ParentForm).Enabled=true;

			}
			
			catch //(System.Threading.ThreadInterruptedException e)
			{
				//MessageBox.Show("Threading Interruption Error " + e.Message.ToString());
			}
		


		}
		private void txtFileInputPopEstUnit()
		{
			//variables used to match up the biosum table columns with the 
			//fiadb txtfile table column names
			int intFIADBColumn=0;
			int intBiosumColumn=0;

			string strValue="";
			string strFields="";
			string strValues="";
			int intAddedRows=0;
			bool bAddRow=true;
			int x=0;
			System.Data.DataTable p_dtChanges;
			System.Data.DataRow[] p_dtNewRows;
			string[,] strList;
			const int CN = 0;
			string strCn="";
			int[] intCol;
			

			try
			{
				//instatiate the oledb data access class
				this.m_ado = new ado_data_access();

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", 6);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Creating Datasource Links");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Text", "Adding Pop Estimation Unit Records");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);

                    
				//get all the project datasources
				FIA_Biosum_Manager.Datasource p_datasource = new Datasource(((frmDialog)this.ParentForm).m_frmMain.frmProject.uc_project1.txtRootDirectory.Text.ToString().Trim());

				//create a temporary mdb file with links to all the project tables
				this.m_strTempMDBFile = p_datasource.CreateMDBAndTableDataSourceLinks();

				//get a connection string for the temp mdb file
				this.m_strTempMDBFileConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"","");

				//get the plot, cond, and tree table names
				m_strPopEstUnitTable = p_datasource.getValidDataSourceTableName("POPULATION ESTIMATION UNIT");

				//remove any entries that have a biosum_status_code of 9
				this.m_ado.SqlNonQuery(this.m_strTempMDBFileConn,"DELETE FROM " + this.m_strPopEstUnitTable + "  WHERE biosum_status_cd = 9;");

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 1);
   					
				//create a new connection
				this.m_connTempMDBFile = new System.Data.OleDb.OleDbConnection();

				//open the connection to the temp mdb file 
				this.m_ado.OpenConnection(this.m_strTempMDBFileConn,ref this.m_connTempMDBFile);

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dtPopEstUnitSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strPopEstUnitTable);

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 2);
			

				//close the connection to the temp mdb file
				this.m_connTempMDBFile.Close();

				/*****************************************************************
				 **create the table structure of the pop eval table and give it 
				 **the name of popestunit_input
				 *****************************************************************/
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"popestunit_input",p_dtPopEstUnitSchema,true);
				if (p_dao.m_intError!=0)
				{
					this.m_intError=p_dao.m_intError;
					p_dao=null;
					p_dtPopEstUnitSchema.Dispose();
					return;
				}

				p_dao=null;

				//reopen the connection to the temp mdb file 
				this.m_ado.OpenConnection(this.m_strTempMDBFileConn,ref this.m_connTempMDBFile);
                   
				//initialize the transaction object with the temporary connection
				System.Data.OleDb.OleDbTransaction p_trans = this.m_connTempMDBFile.BeginTransaction();

				FIA_Biosum_Manager.env p_env = new env();
				System.Data.DataRow p_row;
				this.m_ado.m_DataSet = new DataSet("FIADB");
				this.m_ado.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 1);

				//----------------pop stratum DATA---------------//
				//load the text file into an adodot net datasource table
				this.m_ado.ConvertDelimitedTextToDataTable(this.m_ado.m_DataSet,this.txtPopStratum.Text.Trim(),"fiadb_popestunit",",");

				if (this.m_ado.m_intError==0)
				{
					try
					{
						strList = new string[this.m_ado.m_DataSet.Tables["fiadb_popestunit"].Rows.Count,1];
						for (x=0;x<=this.m_ado.m_DataSet.Tables["fiadb_popestunit"].Rows.Count-1;x++)
						{
							strList[x,CN]="";

						}

						if (this.m_ado.m_DataSet.Tables["fiadb_popestunit"].Rows.Count > 0)
						{
								
								 
							this.m_ado.AddSQLQueryToDataSet(this.m_connTempMDBFile,ref this.m_ado.m_OleDbDataAdapter,ref this.m_ado.m_DataSet,ref p_trans, "select * from " + this.m_strPopEstUnitTable,this.m_strPopEstUnitTable);
								
							strFields = "";
							strValues = "";
							//Build the pop eval insert sql
							for (x=0; x<=p_dtPopEstUnitSchema.Rows.Count-1;x++)
							{
								if (strFields.Trim().Length == 0)
								{
									strFields = "(";
								}
								else
								{	
									strFields = strFields + "," ;
								}
								strFields = strFields + p_dtPopEstUnitSchema.Rows[x]["columnname"].ToString().Trim();
								if (strValues.Trim().Length == 0)
								{
									strValues = "(";
								}
								else
								{	
									strValues = strValues + ",";
								}
								strValues = strValues + "?";

							}
							strFields = strFields + ")";
							strValues = strValues + ");";
							//create an insert command 
							this.m_ado.m_OleDbDataAdapter.InsertCommand = this.m_connTempMDBFile.CreateCommand();
							//bind the transaction object to the insert command
							this.m_ado.m_OleDbDataAdapter.InsertCommand.Transaction = p_trans;
							this.m_ado.m_OleDbDataAdapter.InsertCommand.CommandText = 
								"INSERT INTO " + this.m_strPopEstUnitTable + " "  + strFields + " VALUES " + strValues;
							//define field datatypes for the data adapter
							for (x=0; x<=p_dtPopEstUnitSchema.Rows.Count-1;x++)
							{
								strFields=p_dtPopEstUnitSchema.Rows[x]["columnname"].ToString().Trim();
								switch (p_dtPopEstUnitSchema.Rows[x]["datatype"].ToString().Trim())
								{
									case "System.String" :
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.VarWChar,
											0,
											strFields);
										break;
									case "System.Double":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Double,
											0,
											strFields);
										break;
									case "System.Boolean":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Boolean,
											0,
											strFields);
										break;
									case "System.DateTime":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.DBTimeStamp,
											0,
											strFields);
										break;
									case "System.Decimal":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Decimal,
											0,
											strFields);
										break;
									case "System.Int16":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.SmallInt,
											0,
											strFields);
										break;
									case "System.Int32":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Integer,
											0,
											strFields);
										break;
									case "System.Int64":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.BigInt,
											0,
											strFields);
										break;
									case "System.SByte":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.SmallInt,
											0,
											strFields);
										break;
									case "System.Byte":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.TinyInt,
											0,
											strFields);
										break;
									case "System.Single":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Single,
											0,
											strFields);
										break;
									default:
										MessageBox.Show(this.m_ado.m_DataSet.Tables[this.m_strPopEstUnitTable].Columns[x].DataType.FullName.ToString().Trim());
										break;
								}
									
							}
								
							intCol = new int[this.m_ado.m_DataSet.Tables[this.m_strPopEstUnitTable].Columns.Count];

							//match up the biosum columns with the fiadb columns
							for (intBiosumColumn = 0; intBiosumColumn <= this.m_ado.m_DataSet.Tables[this.m_strPopEstUnitTable].Columns.Count-1;intBiosumColumn++)
							{
								intCol[intBiosumColumn]=-1;
								for (intFIADBColumn = 0; intFIADBColumn <= this.m_ado.m_DataSet.Tables["fiadb_popestunit"].Columns.Count-1;intFIADBColumn++)
								{
									if (this.m_ado.m_DataSet.Tables["fiadb_popestunit"].Columns[intFIADBColumn].ColumnName.Trim().ToUpper() == 
										this.m_ado.m_DataSet.Tables[this.m_strPopEstUnitTable].Columns[intBiosumColumn].ColumnName.Trim().ToUpper())
									{
										intCol[intBiosumColumn] = intFIADBColumn;
									}

								}
							}
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Pop Stratum Table: Compiling New Pop Statum Rows");
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
							
							//load up each row in the FIADB plot input table
							for (x = 0; x<=this.m_ado.m_DataSet.Tables["fiadb_popestunit"].Rows.Count-1;x++)
							{
								bAddRow=true;
								//make sure the row is not null values
								if (this.m_ado.m_DataSet.Tables["fiadb_popestunit"].Rows[x][0] != System.DBNull.Value &&
									this.m_ado.m_DataSet.Tables["fiadb_popestunit"].Rows[x][0].ToString().Trim().Length > 0)
								{
									
									p_row = this.m_ado.m_DataSet.Tables[this.m_strPopEstUnitTable].NewRow();
									for (intBiosumColumn = 0; intBiosumColumn <= intCol.Length-1; intBiosumColumn++)
									{
										if (intCol[intBiosumColumn] != -1)
										{
											strValue = this.m_ado.m_DataSet.Tables["fiadb_popestunit"].Rows[x][intCol[intBiosumColumn]].ToString().Trim();
											if (strValue.Trim().Length > 0)
											{
												switch (this.m_ado.m_DataSet.Tables[this.m_strPopEstUnitTable].Columns[intBiosumColumn].DataType.FullName.ToString().Trim())
												{
													case "System.String" :
														p_row[intBiosumColumn] = strValue.Replace("\"","");
														break;
													case "System.Double":
														p_row[intBiosumColumn] = Convert.ToDouble(strValue);
														break;
													case "System.Boolean":
														p_row[intBiosumColumn] = Convert.ToBoolean(strValue);
														break;
													case "System.DateTime":
														p_row[intBiosumColumn] = Convert.ToDateTime(strValue);
														break;
													case "System.Decimal":
														p_row[intBiosumColumn] = Convert.ToDecimal(strValue);
														break;
													case "System.Int16":
														p_row[intBiosumColumn] = Convert.ToInt16(strValue);
														break;
													case "System.Int32":
														p_row[intBiosumColumn] = Convert.ToInt32(strValue);
														break;
													case "System.Int64":
														p_row[intBiosumColumn] = Convert.ToInt64(strValue);
														break;
													case "System.SByte":
														p_row[intBiosumColumn] = Convert.ToSByte(strValue);

														break;
													case "System.Single":
														p_row[intBiosumColumn] = Convert.ToSingle(strValue);
														break;

													default:
														MessageBox.Show(this.m_ado.m_DataSet.Tables[this.m_strPopEstUnitTable].Columns[intBiosumColumn].DataType.FullName.ToString().Trim());
														break;
												}
											}
										}
										System.Windows.Forms.Application.DoEvents();
										if (this.m_frmTherm.AbortProcess == true) break;
									}
									if (this.m_frmTherm.AbortProcess == false)
									{
											
										//get the plt_cn value
										strCn = this.m_ado.m_DataSet.Tables["fiadb_popestunit"].Rows[x]["cn"].ToString().Trim();
										strCn = strCn.Replace("\"","");

										//see if user has a list to filter the plots to process
										if (this.m_strPlotIdList.Trim().Length > 0 )
										{
											if (this.m_strPlotIdList.IndexOf("'" + strCn.Trim() + "'") < 0)
											{
												bAddRow=false;
											}
										}
										
										//check rscd and evalid
										if (m_ado.m_DataSet.Tables["fiadb_popestunit"].Rows[x]["rscd"] != System.DBNull.Value &&
											m_ado.m_DataSet.Tables["fiadb_popestunit"].Rows[x]["evalid"] != System.DBNull.Value)
										{
											if (Convert.ToString(m_ado.m_DataSet.Tables["fiadb_popestunit"].Rows[x]["rscd"]).Trim() != this.m_strCurrFIADBRsCd.Trim() ||
												Convert.ToString(m_ado.m_DataSet.Tables["fiadb_popestunit"].Rows[x]["evalid"]).Trim() != this.m_strCurrFIADBEvalId.Trim())
											{
												bAddRow=false;
											}


										}
										if (bAddRow==true)
										{
											if (p_row["biosum_status_cd"] == System.DBNull.Value) 
												p_row["biosum_status_cd"] = 9;		
											strList[intAddedRows,CN]=strCn.Trim();
											this.m_ado.m_DataSet.Tables[this.m_strPopEstUnitTable].Rows.Add(p_row);
											intAddedRows++;
											
										}
									}
								}
							
							}
						}
						else
						{
							MessageBox.Show("!!No Pop Eval Records In The FIADB Pop Eval Input Table!!");
						}
					}
					catch (Exception caught)
					{
						this.m_intError=-1;
						MessageBox.Show(caught.Message);
					}
					//keep processing the records if there was no error and there were plot records added
					if (this.m_intError==0 && intAddedRows > 0 && this.m_frmTherm.AbortProcess == false)
					{
						try
						{
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Pop Eval Table: Finishing New Pop Eval Rows...Stand By");
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
							
							p_dtChanges = this.m_ado.m_DataSet.Tables[this.m_strPopEstUnitTable].GetChanges();
								
							p_dtNewRows = p_dtChanges.Select(null,null, DataViewRowState.Added);
							if (p_dtChanges.HasErrors)
							{
								this.m_ado.m_DataSet.Tables[this.m_strPopEstUnitTable].RejectChanges();
								this.m_intError=-1;
							}
							else
							{
								this.m_ado.m_OleDbDataAdapter.Update(p_dtNewRows);
								this.m_ado.m_DataSet.Tables[this.m_strPopEstUnitTable].AcceptChanges();
							}
						}
						catch (Exception caught)
						{
							this.m_intError=-1;
							this.m_ado.m_DataSet.Tables[this.m_strPopEstUnitTable].RejectChanges();
							//rollback the transaction to the original records 
							p_trans.Rollback();
							MessageBox.Show(caught.Message);
						}
					}
				}

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 4);
				
				
				if (this.m_frmTherm.AbortProcess==false && this.m_intError == 0 && (intAddedRows > 0))
				{
					try
					{
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 2);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Committing The Data...Stand by");
                        frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
						p_trans.Commit(); 

					}
					catch //(Exception caught)
					{
						p_trans.Rollback();
						this.m_intError=-1;
					}

				}
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 5);
				if (this.m_frmTherm.AbortProcess == true &&  (intAddedRows > 0))
				{
					p_trans.Rollback();
				
				}
				else
				{
					if (this.m_frmTherm.AbortProcess==false && this.m_ado.m_intError==0 && this.m_intError == 0 && (intAddedRows > 0))
					{
						p_trans = null;

						if (this.m_intError==0)
						{
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Done");
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", (int)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1, "Maximum", false));
						}
						else
						{
							//error occurred in the updatecolumns so delete the records
							this.m_strSQL = "DELETE FROM " + this.m_strPopEstUnitTable + " WHERE biosum_status_cd = 9;";
							this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
							MessageBox.Show("!!Error Occured Adding Pop Eval Records: 0 Records Added!!","Add Pop Estimation Unit Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
						}
					}

				}
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", (int)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1, "Maximum", false));
                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form)m_frmTherm, "Close");
				this.m_frmTherm = null;
				p_dtChanges = null;
				p_dtNewRows = null;
				this.m_ado.m_DataSet.Clear();
				this.m_ado.m_DataSet.Dispose();
				this.m_ado = null;
				((frmDialog)this.ParentForm).Enabled=true;

			}
			
			catch //(System.Threading.ThreadInterruptedException e)
			{
				//MessageBox.Show("Threading Interruption Error " + e.Message.ToString());
			}
		


		}

		private void txtFileInputPopFiles()
		{
			//variables used to match up the biosum table columns with the 
			//fiadb txtfile table column names

			int intFIADBColumn=0;
			int intBiosumColumn=0;
			string strTable="";
			string strValue="";
			string strFields="";
			string strValues="";
			int intAddedRows=0;
			bool bAddRow=true;
			int x=0;
			System.Data.DataTable p_dtChanges;
			System.Data.DataRow[] p_dtNewRows;
			string[,] strList;
			const int CN = 0;
			string strCn="";
			int[] intCol;
		    string  strRound;

			this.m_bLoadStateCountyList=true;
			this.m_bLoadStateCountyPlotList=true;
						
			

			try
			{
				this.m_intError=0;
				//get all the project datasources
				//FIA_Biosum_Manager.Datasource m_oDataSource = new Datasource(((frmDialog)this.ParentForm).m_frmMain.frmProject.uc_project1.txtRootDirectory.Text.ToString().Trim());
				//get the table name
				strTable = this.m_oDatasource.getValidDataSourceTableName(this.m_strTableType);


				//instatiate the oledb data access class
				this.m_ado = new ado_data_access();

				

				//create a temporary mdb file with links to all the project tables
				this.m_strTempMDBFile = m_oDatasource.CreateMDBAndTableDataSourceLinks();

				//get a connection string for the temp mdb file
				this.m_strTempMDBFileConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"","");

				//remove any entries that have a biosum_status_code of 9
				this.m_ado.SqlNonQuery(this.m_strTempMDBFileConn,"DELETE FROM " + strTable + "  WHERE biosum_status_cd = 9;");

				if (this.m_strTableType.Trim().ToUpper() == "POPULATION EVALUATION")
                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 1);
   					
				//create a new connection
				this.m_connTempMDBFile = new System.Data.OleDb.OleDbConnection();

				//open the connection to the temp mdb file 
				this.m_ado.OpenConnection(this.m_strTempMDBFileConn,ref this.m_connTempMDBFile);

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dtPopTableSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + strTable);

				if (this.m_strTableType.Trim().ToUpper() == "POPULATION EVALUATION")
                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 2);
			

				//close the connection to the temp mdb file
				this.m_connTempMDBFile.Close();

				/*****************************************************************
				 **create the table structure of the pop eval table and give it 
				 **the name of poptable_input
				 *****************************************************************/
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"poptable_input",p_dtPopTableSchema,true);
				if (p_dao.m_intError!=0)
				{
					this.m_intError=p_dao.m_intError;
					p_dao=null;
					p_dtPopTableSchema.Dispose();
					return;
				}
				p_dao=null;

				//reopen the connection to the temp mdb file 
				this.m_ado.OpenConnection(this.m_strTempMDBFileConn,ref this.m_connTempMDBFile);
                   
				//initialize the transaction object with the temporary connection
				System.Data.OleDb.OleDbTransaction p_trans = this.m_connTempMDBFile.BeginTransaction();

				FIA_Biosum_Manager.env p_env = new env();
				System.Data.DataRow p_row;
				this.m_ado.m_DataSet = new DataSet("FIADB");
				this.m_ado.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
				

				//----------------pop stratum DATA---------------//
				//load the text file into an adodot net datasource table
				this.m_ado.ConvertDelimitedTextToDataTable(this.m_ado.m_DataSet,m_strCurrentTxtInputFile.Trim(),"fiadb_poptable",",");

				if (this.m_ado.m_intError==0)
				{
					try
					{
						strList = new string[this.m_ado.m_DataSet.Tables["fiadb_poptable"].Rows.Count,1];
						for (x=0;x<=this.m_ado.m_DataSet.Tables["fiadb_poptable"].Rows.Count-1;x++)
						{
							strList[x,CN]="";

						}

						if (this.m_ado.m_DataSet.Tables["fiadb_poptable"].Rows.Count > 0)
						{
								
								 
							this.m_ado.AddSQLQueryToDataSet(this.m_connTempMDBFile,ref this.m_ado.m_OleDbDataAdapter,ref this.m_ado.m_DataSet,ref p_trans, "select * from " + strTable,strTable);
								
							strFields = "";
							strValues = "";
							//Build the pop eval insert sql
							for (x=0; x<=p_dtPopTableSchema.Rows.Count-1;x++)
							{
								if (strFields.Trim().Length == 0)
								{
									strFields = "(";
								}
								else
								{	
									strFields = strFields + "," ;
								}
								strFields = strFields + p_dtPopTableSchema.Rows[x]["columnname"].ToString().Trim();
								if (strValues.Trim().Length == 0)
								{
									strValues = "(";
								}
								else
								{	
									strValues = strValues + ",";
								}
								strValues = strValues + "?";

							}
							strFields = strFields + ")";
							strValues = strValues + ");";
							//create an insert command 
							this.m_ado.m_OleDbDataAdapter.InsertCommand = this.m_connTempMDBFile.CreateCommand();
							//bind the transaction object to the insert command
							this.m_ado.m_OleDbDataAdapter.InsertCommand.Transaction = p_trans;
							this.m_ado.m_OleDbDataAdapter.InsertCommand.CommandText = 
								"INSERT INTO " + strTable + " "  + strFields + " VALUES " + strValues;
							//define field datatypes for the data adapter
							for (x=0; x<=p_dtPopTableSchema.Rows.Count-1;x++)
							{
								
								strFields=p_dtPopTableSchema.Rows[x]["columnname"].ToString().Trim();
								switch (p_dtPopTableSchema.Rows[x]["datatype"].ToString().Trim())
								{
									case "System.String" :
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.VarWChar,
											0,
											strFields);
										break;
									case "System.Double":
										
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Double,
											0,
											strFields);
										break;
									case "System.Boolean":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Boolean,
											0,
											strFields);
										break;
									case "System.DateTime":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.DBTimeStamp,
											0,
											strFields);
										break;
									case "System.Decimal":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Decimal,
											0,
											strFields);
										break;
									case "System.Int16":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.SmallInt,
											0,
											strFields);
										break;
									case "System.Int32":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Integer,
											0,
											strFields);
										break;
									case "System.Int64":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.BigInt,
											0,
											strFields);
										break;
									case "System.SByte":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.SmallInt,
											0,
											strFields);
										break;
									case "System.Byte":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.TinyInt,
											0,
											strFields);
										break;
									case "System.Single":
										this.m_ado.m_OleDbDataAdapter.InsertCommand.Parameters.Add
											(strFields, 
											System.Data.OleDb.OleDbType.Single,
											0,
											strFields);
										break;
									default:
										MessageBox.Show(this.m_ado.m_DataSet.Tables[strTable].Columns[x].DataType.FullName.ToString().Trim());
										break;
								}
									
							}
								
							intCol = new int[this.m_ado.m_DataSet.Tables[strTable].Columns.Count];

							//match up the biosum columns with the fiadb columns
							for (intBiosumColumn = 0; intBiosumColumn <= this.m_ado.m_DataSet.Tables[strTable].Columns.Count-1;intBiosumColumn++)
							{
								intCol[intBiosumColumn]=-1;
								for (intFIADBColumn = 0; intFIADBColumn <= this.m_ado.m_DataSet.Tables["fiadb_poptable"].Columns.Count-1;intFIADBColumn++)
								{
									if (this.m_ado.m_DataSet.Tables["fiadb_poptable"].Columns[intFIADBColumn].ColumnName.Trim().ToUpper() == 
										this.m_ado.m_DataSet.Tables[strTable].Columns[intBiosumColumn].ColumnName.Trim().ToUpper())
									{
										intCol[intBiosumColumn] = intFIADBColumn;
									}

								}
							}

                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", (int)this.m_ado.m_DataSet.Tables["fiadb_poptable"].Rows.Count - 1);
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", strTable + " Table: Compiling New Rows");
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

							//load up each row in the FIADB plot input table
							for (x = 0; x<=this.m_ado.m_DataSet.Tables["fiadb_poptable"].Rows.Count-1;x++)
							{
                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", (int)x);
								bAddRow=true;
								//make sure the row is not null values
								if (this.m_ado.m_DataSet.Tables["fiadb_poptable"].Rows[x][0] != System.DBNull.Value &&
									this.m_ado.m_DataSet.Tables["fiadb_poptable"].Rows[x][0].ToString().Trim().Length > 0)
								{
									
									p_row = this.m_ado.m_DataSet.Tables[strTable].NewRow();
									for (intBiosumColumn = 0; intBiosumColumn <= intCol.Length-1; intBiosumColumn++)
									{
										if (intCol[intBiosumColumn] != -1)
										{
											strRound="";
											switch (this.m_ado.m_DataSet.Tables[strTable].Columns[intBiosumColumn].DataType.FullName.ToString().Trim())
											{
												case "System.Double":
													strRound="double";
													break;
												case "System.Decimal":
													strRound="decimal";
													break;
												
											}
											if (this.m_ado.m_DataSet.Tables["fiadb_poptable"].Rows[x][intCol[intBiosumColumn]]!= System.DBNull.Value)
											{
												if (strRound.Length > 0)
												{
													/************************************************************************************
													 *the code below fixes a scientfic notation error from the CSV files that are input.*
													 *If the data type is decimal or double the input data can be represented with      *
													 *scientific notation. If the scientific notation is defined as E00 then an error   *
													 *exception occurs. The fix is to replace E00 with nothing.                         *
													 ************************************************************************************/
													this.m_ado.m_DataSet.Tables["fiadb_poptable"].Rows[x][intCol[intBiosumColumn]]=this.m_ado.m_DataSet.Tables["fiadb_poptable"].Rows[x][intCol[intBiosumColumn]].ToString().Replace("E00","");
													if (strRound=="double")
													{
														strValue = Convert.ToString(Math.Round(Convert.ToDouble(this.m_ado.m_DataSet.Tables["fiadb_poptable"].Rows[x][intCol[intBiosumColumn]]),14)).Trim();
													}
													else
													{
														strValue = Convert.ToString(Math.Round(Convert.ToDecimal(this.m_ado.m_DataSet.Tables["fiadb_poptable"].Rows[x][intCol[intBiosumColumn]]),14)).Trim();
													}

												}
												else
												{
													strValue = this.m_ado.m_DataSet.Tables["fiadb_poptable"].Rows[x][intCol[intBiosumColumn]].ToString().Trim();
												}
											}
											else strValue="";
											if (strValue.Trim().Length > 0)
											{
												switch (this.m_ado.m_DataSet.Tables[strTable].Columns[intBiosumColumn].DataType.FullName.ToString().Trim())
												{
													case "System.String" :
														p_row[intBiosumColumn] = strValue.Replace("\"","");
														break;
													case "System.Double":
														p_row[intBiosumColumn] = Convert.ToDouble(strValue);
														break;
													case "System.Boolean":
														p_row[intBiosumColumn] = Convert.ToBoolean(strValue);
														break;
													case "System.DateTime":
														p_row[intBiosumColumn] = Convert.ToDateTime(strValue);
														break;
													case "System.Decimal":
														p_row[intBiosumColumn] = Convert.ToDecimal(strValue);
														break;
													case "System.Int16":
														p_row[intBiosumColumn] = Convert.ToInt16(strValue);
														break;
													case "System.Int32":
														p_row[intBiosumColumn] = Convert.ToInt32(strValue);
														break;
													case "System.Int64":
														p_row[intBiosumColumn] = Convert.ToInt64(strValue);
														break;
													case "System.SByte":
														p_row[intBiosumColumn] = Convert.ToSByte(strValue);

														break;
													case "System.Single":
														p_row[intBiosumColumn] = Convert.ToSingle(strValue);
														break;

													default:
														MessageBox.Show(this.m_ado.m_DataSet.Tables[strTable].Columns[intBiosumColumn].DataType.FullName.ToString().Trim());
														break;
												}
											}
										}
										System.Windows.Forms.Application.DoEvents();
										if (this.m_frmTherm.AbortProcess == true) break;
									}
									if (this.m_frmTherm.AbortProcess == false)
									{
											
										//get the plt_cn value
										strCn = this.m_ado.m_DataSet.Tables["fiadb_poptable"].Rows[x]["cn"].ToString().Trim();
										strCn = strCn.Replace("\"","");

										//see if user has a list to filter the plots to process
										if (this.m_strPlotIdList.Trim().Length > 0 )
										{
											if (this.m_strPlotIdList.IndexOf("'" + strCn.Trim() + "'") < 0)
											{
												bAddRow=false;
											}
										}
										
										if (this.m_strTableType.Trim().ToUpper() != "POPULATION EVALUATION")
										{
											//check rscd and evalid
											if (m_ado.m_DataSet.Tables["fiadb_poptable"].Rows[x]["rscd"] != System.DBNull.Value &&
												m_ado.m_DataSet.Tables["fiadb_poptable"].Rows[x]["evalid"] != System.DBNull.Value)
											{
												if (Convert.ToString(m_ado.m_DataSet.Tables["fiadb_poptable"].Rows[x]["rscd"]).Trim() != this.m_strCurrFIADBRsCd.Trim() ||
													Convert.ToString(m_ado.m_DataSet.Tables["fiadb_poptable"].Rows[x]["evalid"]).Trim() != this.m_strCurrFIADBEvalId.Trim())
												{
													bAddRow=false;
												}


											}
										}
										if (bAddRow==true)
										{
											if (p_row["biosum_status_cd"] == System.DBNull.Value) 
												p_row["biosum_status_cd"] = 9;		
											strList[intAddedRows,CN]=strCn.Trim();
											this.m_ado.m_DataSet.Tables[strTable].Rows.Add(p_row);
											intAddedRows++;
											
										}
										else
										{
											p_row.Delete();
										}
									}
								}
							
							}
						}
						else
						{
							MessageBox.Show("!!No Pop Eval Records In The FIADB Pop Eval Input Table!!");
						}
					}
					catch (Exception caught)
					{
						this.m_intError=-1;
						MessageBox.Show(caught.Message);
					}
					//keep processing the records if there was no error and there were plot records added
					if (this.m_intError==0 && intAddedRows > 0 && this.m_frmTherm.AbortProcess == false)
					{
						try
						{

							//if (this.m_strTableType.Trim().ToUpper() == "POPULATION EVALUATION")
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", strTable + " Table: Finishing New Rows...Stand By");
								
							//else
							//	this.m_frmTherm.lblMsg2.Text=strTable + " Table: Finishing New Rows...Stand By";

                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
							p_dtChanges = this.m_ado.m_DataSet.Tables[strTable].GetChanges();
								
							p_dtNewRows = p_dtChanges.Select(null,null, DataViewRowState.Added);
							if (p_dtChanges.HasErrors)
							{
								this.m_ado.m_DataSet.Tables[strTable].RejectChanges();
								this.m_intError=-1;
							}
							else
							{
								this.m_ado.m_OleDbDataAdapter.Update(p_dtNewRows);
								this.m_ado.m_DataSet.Tables[strTable].AcceptChanges();
							}
						}
						catch (Exception caught)
						{
							this.m_intError=-1;
							this.m_ado.m_DataSet.Tables[strTable].RejectChanges();
							//rollback the transaction to the original records 
							p_trans.Rollback();
							MessageBox.Show(caught.Message);
						}
					}
				}
				
				if (this.m_frmTherm.AbortProcess==false && this.m_intError == 0 && (intAddedRows > 0))
				{
					try
					{
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Committing The Data...Stand by");

                        frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
						p_trans.Commit(); 

					}
					catch //(Exception caught)
					{
						p_trans.Rollback();
						this.m_intError=-1;
					}

				}
				
				if (this.m_frmTherm.AbortProcess == true &&  (intAddedRows > 0))
				{
					p_trans.Rollback();
				
				}
				else
				{
					if (this.m_frmTherm.AbortProcess==false && this.m_ado.m_intError==0 && this.m_intError == 0 && (intAddedRows > 0))
					{
						p_trans = null;
						
						if (this.m_intError==0)
						{
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Done");
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

						}
						else
						{
							//error occurred in the updatecolumns so delete the records
							this.m_strSQL = "DELETE FROM " + strTable + " WHERE biosum_status_cd = 9;";
							this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
							MessageBox.Show("!!Error Occured Adding Pop Eval Records: 0 Records Added!!","Add Pop Estimation Unit Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
						}
					}

				}
				
				p_dtChanges = null;
				p_dtNewRows = null;
				this.m_ado.m_DataSet.Clear();
				this.m_ado.m_DataSet.Dispose();
				this.m_ado = null;

				((frmDialog)this.ParentForm).Enabled=true;
				

			}
			
			catch //(System.Threading.ThreadInterruptedException e)
			{
				//MessageBox.Show("Threading Interruption Error " + e.Message.ToString());
			}
		


		}

		private string CreateBiosumPlotId(System.Data.DataRow p_dr)
		{
			string strBiosumPlotId="";
			string strInvId ="";
			string strStateCd = "";
			string strCycle="";
			string strSubCycle="";
			string strCountyCd="";
			string strPlot="";
			string strForestBlm="";

			//inventory id
			if (this.rdoFIADB.Checked==true)
			{
				strBiosumPlotId = "1";
				if (p_dr["measyear"] != System.DBNull.Value &&
					p_dr["measyear"].ToString().Trim().Length > 0)
				{
					strInvId = p_dr["measyear"].ToString().Trim();
				}
				else
				{
					strInvId = "9999";
				}
			}
			else if (this.rdoIDB.Checked==true)
			{
				strBiosumPlotId = "2";
				//idb inventory
				if (strInvId.Trim().Length == 0) strInvId = "9999";
				
               
			}

			strBiosumPlotId = strBiosumPlotId + strInvId;


			
			//state
			if (p_dr["statecd"] != System.DBNull.Value &&
				p_dr["statecd"].ToString().Trim().Length > 0)
			{
				strStateCd= p_dr["statecd"].ToString().Trim();
			}

			switch (strStateCd.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "99";
					break;
				case 1:
					strBiosumPlotId =  strBiosumPlotId + "0" + strStateCd.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strStateCd.Trim();
					break;
			}

			//cycle
			if (this.rdoFIADB.Checked==true)
			{
				if (p_dr["cycle"] != System.DBNull.Value &&
					p_dr["cycle"].ToString().Trim().Length > 0)
				{
					strCycle = p_dr["cycle"].ToString().Trim();
				}
				
			}
			switch (strCycle.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "00";
					break;
				case 1:
					strBiosumPlotId =  strBiosumPlotId + "0" + strCycle.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strCycle.Trim();
					break;
			}

			//subcycle
			if (this.rdoFIADB.Checked==true)
			{
				if (p_dr["subcycle"] != System.DBNull.Value &&
					p_dr["subcycle"].ToString().Trim().Length > 0)
				{
					strSubCycle = p_dr["subcycle"].ToString().Trim();
				}
				
			}
			switch (strSubCycle.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "00";
					break;
				case 1:
					strBiosumPlotId =  strBiosumPlotId + "0" + strSubCycle.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strSubCycle.Trim();
					break;
			}


			//countycode

			if (p_dr["countycd"] != System.DBNull.Value &&
				p_dr["countycd"].ToString().Trim().Length > 0)
			{
				strCountyCd = p_dr["countycd"].ToString().Trim();
			}

			switch (strCountyCd.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "999";
					break;
				case 1:
					strBiosumPlotId = strBiosumPlotId + "00" + strCountyCd.Trim();
					break;
				case 2:
					strBiosumPlotId = strBiosumPlotId + "0" + strCountyCd.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strCountyCd.Trim();
					break;
			}

			//plot
			if (p_dr["plot"] != System.DBNull.Value &&
				p_dr["plot"].ToString().Trim().Length > 0)
			{
				strPlot = p_dr["plot"].ToString().Trim();
			}


			switch (strPlot.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "9999999";
					break;
				case 1:
					strBiosumPlotId = strBiosumPlotId + "000000" + strPlot.Trim();
					break;
				case 2:
					strBiosumPlotId = strBiosumPlotId + "00000" + strPlot.Trim();
					break;
				case 3:
					strBiosumPlotId = strBiosumPlotId + "0000" + strPlot.Trim();
					break;
				case 4:
					strBiosumPlotId = strBiosumPlotId + "000" + strPlot.Trim();
					break;
				case 5:
					strBiosumPlotId = strBiosumPlotId + "00" + strPlot.Trim();
					break;
				case 6:
					strBiosumPlotId = strBiosumPlotId + "0" + strPlot.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strPlot.Trim();
					break;
			}


			
			//forest or blm district - need value for pnw idb unique key value
			if (this.rdoFIADB.Checked==true)
			{
				strForestBlm="000";
			}
			else if (this.rdoIDB.Checked==true)
			{
				if (p_dr["forest_or_blm_district"] != System.DBNull.Value &&
					p_dr["forest_or_blm_district"].ToString().Trim().Length > 0)
				{
					strForestBlm = p_dr["forest_or_blm_district"].ToString().Trim();
				}
              
			}


			switch (strForestBlm.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "999";
					break;
				case 1:
					strBiosumPlotId = strBiosumPlotId + "00" + strForestBlm.Trim();
					break;
				case 2:
					strBiosumPlotId = strBiosumPlotId + "0" + strForestBlm.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strForestBlm.Trim();
					break;
			}
			return strBiosumPlotId;
		}
		private string CreateBiosumPlotId(System.Data.OleDb.OleDbDataReader p_dr)
		{
			string strBiosumPlotId="";
			string strInvId ="";
			string strStateCd = "";
			string strCycle="";
			string strSubCycle="";
			string strCountyCd="";
			string strPlot="";
			string strForestBlm="";

			
			//inventory id
			if (this.rdoFIADB.Checked==true)
			{
				strBiosumPlotId = "1";
				if (p_dr["measyear"] != System.DBNull.Value &&
					p_dr["measyear"].ToString().Trim().Length > 0)
				{
					strInvId = p_dr["measyear"].ToString().Trim();
				}
				else
				{
					strInvId = "9999";
				}
			}
			else if (this.rdoIDB.Checked==true)
			{
				strBiosumPlotId = "2";
				//idb inventory
				if (strInvId.Trim().Length == 0) strInvId = "9999";
				
               
			}

			strBiosumPlotId = strBiosumPlotId + strInvId;


			
			//state
			if (p_dr["statecd"] != System.DBNull.Value &&
				p_dr["statecd"].ToString().Trim().Length > 0)
			{
				strStateCd= p_dr["statecd"].ToString().Trim();
			}

			switch (strStateCd.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "99";
					break;
				case 1:
					strBiosumPlotId =  strBiosumPlotId + "0" + strStateCd.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strStateCd.Trim();
					break;
			}

			//cycle
			if (this.rdoFIADB.Checked==true)
			{
				if (p_dr["cycle"] != System.DBNull.Value &&
					p_dr["cycle"].ToString().Trim().Length > 0)
				{
					strCycle = p_dr["cycle"].ToString().Trim();
				}
				
			}
			switch (strCycle.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "00";
					break;
				case 1:
					strBiosumPlotId =  strBiosumPlotId + "0" + strCycle.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strCycle.Trim();
					break;
			}

			//subcycle
			if (this.rdoFIADB.Checked==true)
			{
				if (p_dr["subcycle"] != System.DBNull.Value &&
					p_dr["subcycle"].ToString().Trim().Length > 0)
				{
					strSubCycle = p_dr["subcycle"].ToString().Trim();
				}
				
			}
			switch (strSubCycle.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "00";
					break;
				case 1:
					strBiosumPlotId =  strBiosumPlotId + "0" + strSubCycle.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strSubCycle.Trim();
					break;
			}


			//countycode

			if (p_dr["countycd"] != System.DBNull.Value &&
				p_dr["countycd"].ToString().Trim().Length > 0)
			{
				strCountyCd = p_dr["countycd"].ToString().Trim();
			}

			switch (strCountyCd.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "999";
					break;
				case 1:
					strBiosumPlotId = strBiosumPlotId + "00" + strCountyCd.Trim();
					break;
				case 2:
					strBiosumPlotId = strBiosumPlotId + "0" + strCountyCd.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strCountyCd.Trim();
					break;
			}

			//plot
			if (p_dr["plot"] != System.DBNull.Value &&
				p_dr["plot"].ToString().Trim().Length > 0)
			{
				strPlot = p_dr["plot"].ToString().Trim();
			}


			switch (strPlot.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "9999999";
					break;
				case 1:
					strBiosumPlotId = strBiosumPlotId + "000000" + strPlot.Trim();
					break;
				case 2:
					strBiosumPlotId = strBiosumPlotId + "00000" + strPlot.Trim();
					break;
				case 3:
					strBiosumPlotId = strBiosumPlotId + "0000" + strPlot.Trim();
					break;
				case 4:
					strBiosumPlotId = strBiosumPlotId + "000" + strPlot.Trim();
					break;
				case 5:
					strBiosumPlotId = strBiosumPlotId + "00" + strPlot.Trim();
					break;
				case 6:
					strBiosumPlotId = strBiosumPlotId + "0" + strPlot.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strPlot.Trim();
					break;
			}


			
			//forest or blm district - need value for pnw idb unique key value
			if (this.rdoFIADB.Checked==true)
			{
				strForestBlm="000";
			}
			else if (this.rdoIDB.Checked==true)
			{
				if (p_dr["forest_or_blm_district"] != System.DBNull.Value &&
					p_dr["forest_or_blm_district"].ToString().Trim().Length > 0)
				{
					strForestBlm = p_dr["forest_or_blm_district"].ToString().Trim();
				}
              
			}


			switch (strForestBlm.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "999";
					break;
				case 1:
					strBiosumPlotId = strBiosumPlotId + "00" + strForestBlm.Trim();
					break;
				case 2:
					strBiosumPlotId = strBiosumPlotId + "0" + strForestBlm.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strForestBlm.Trim();
					break;
			}
			return strBiosumPlotId;
		}
        private void CleanupThread()
        {
           // ((frmDialog)this.ParentForm).m_frmMain.Visible = true;
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog,"Visible",true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);
            
        }
        private void ThreadCleanUp()
        {
            try
            {
               // ((frmDialog)this.ParentForm).m_frmMain.Visible = true;
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Visible", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);

                if (this.m_frmTherm != null)
                {
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form)m_frmTherm, "Close");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form)m_frmTherm, "Dispose");

                    this.m_frmTherm = null;
                }
                
            }
            catch
            {
            }

        }
        private void CancelThread()
        {
            bool bAbort = frmMain.g_oDelegate.AbortProcessing("FIA Biosum", "Do you wish to cancel processing (Y/N)?");
            if (bAbort)
            {
                if (frmMain.g_oDelegate.m_oThread.IsAlive)
                {
                    frmMain.g_oDelegate.m_oThread.Join();
                }
                frmMain.g_oDelegate.StopThread();
                CleanupThread();
            }
        }
        private void CalculateAdjustments_Process()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.CalculateAdjustments_Process\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }

            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
           
            string strFields = "";

            int x = 0;
            int y = 0;
            string strCol = "";
            string str = "";
            string str2 = "";

            string strSourceTableName = "";
            string strDestTableLinkName = "";
            string strFIADBDbFile = "";
            //string strMsg="";

            m_intAddedPlotRows = 0;
            m_intAddedCondRows = 0;
            m_intAddedTreeRows = 0;
            m_intAddedSiteTreeRows = 0;

            this.m_intError = 0;
            System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();

            //-----------PREPARATION FOR CALCULATING ADJUSTMENTS---------//

            try
            {
                //instatiate the oledb data access class
                this.m_ado = new ado_data_access();

                //progress bar 1: single process
                this.SetThermValue(m_frmTherm.progressBar1, "Maximum", 100);
                this.SetThermValue(m_frmTherm.progressBar1, "Minimum", 0);
                this.SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                this.SetLabelValue(m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);

                //progress bar 2: overall progress
                this.SetThermValue(m_frmTherm.progressBar2, "Maximum", 100);
                this.SetThermValue(m_frmTherm.progressBar2, "Minimum", 0);
                this.SetThermValue(m_frmTherm.progressBar2, "Value", 0);
                this.SetLabelValue(m_frmTherm.lblMsg2, "Text", "Overall Progress");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);

                

                //open the FIADB ACCESS DbFile
                m_ado.OpenConnection(m_ado.getMDBConnString(txtMDBFiadbInputFile.Text.Trim(), "", ""),ref oConn);

                
                m_intError = m_ado.m_intError;

                if (this.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    this.SetLabelValue(m_frmTherm.lblMsg, "Text", "Drop Work Tables");
                    SetThermValue(m_frmTherm.progressBar1, "Value", 10);
                    if (m_ado.TableExist(oConn, "BIOSUM_PLOT"))
                        m_ado.SqlNonQuery(oConn, "DROP TABLE BIOSUM_PLOT");
                    SetThermValue(m_frmTherm.progressBar1, "Value", 20);
                    if (m_ado.TableExist(oConn, "BIOSUM_COND"))
                        m_ado.SqlNonQuery(oConn, "DROP TABLE BIOSUM_COND");
                    SetThermValue(m_frmTherm.progressBar1, "Value", 30);
                    if (m_ado.TableExist(oConn, frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName))
                        m_ado.SqlNonQuery(oConn, "DROP TABLE " + frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName);
                     SetThermValue(m_frmTherm.progressBar1, "Value", 40);
                    if (m_ado.TableExist(oConn, "BIOSUM_PPSA"))
                        m_ado.SqlNonQuery(oConn, "DROP TABLE BIOSUM_PPSA");
                     SetThermValue(m_frmTherm.progressBar1, "Value", 50);
                    if (m_ado.TableExist(oConn, "BIOSUM_EUS_TEMP"))
                        m_ado.SqlNonQuery(oConn, "DROP TABLE BIOSUM_EUS_TEMP");
                     SetThermValue(m_frmTherm.progressBar1, "Value", 60);
                    if (m_ado.TableExist(oConn, "BIOSUM_PPSA_DENIED_ACCESS"))
                        m_ado.SqlNonQuery(oConn, "DROP TABLE BIOSUM_PPSA_DENIED_ACCESS");
                     SetThermValue(m_frmTherm.progressBar1, "Value", 70);
                    if (m_ado.TableExist(oConn, "BIOSUM_PPSA_TEMP"))
                        m_ado.SqlNonQuery(oConn, "DROP TABLE BIOSUM_PPSA_TEMP");
                     SetThermValue(m_frmTherm.progressBar1, "Value", 80);
                    if (m_ado.TableExist(oConn, "BIOSUM_EUS_TEMP"))
                        m_ado.SqlNonQuery(oConn, "DROP TABLE BIOSUM_EUS_TEMP");
                     SetThermValue(m_frmTherm.progressBar1, "Value", 90);
                    if (m_ado.TableExist(oConn, "BIOSUM_EUS_ACCESS"))
                        m_ado.SqlNonQuery(oConn, "DROP TABLE BIOSUM_EUS_ACCESS");

                
                    SetThermValue(m_frmTherm.progressBar1, "Value", 100);
                    System.Threading.Thread.Sleep(2000);
                    SetThermValue(m_frmTherm.progressBar2, "Value", 30);
                    string[] strSql = Queries.FIAPlot.FIADBPlotInput_CalculateAdjustmentFactorsSQL(
                        "POP_PLOT_STRATUM_ASSGN",
                        "POP_ESTN_UNIT",
                        "POP_STRATUM",
                        "POP_EVAL",
                        "PLOT",
                        "COND",
                         m_strCurrFIADBRsCd,
                         m_strCurrFIADBEvalId,
                         frmMain.g_oDelegate.GetControlPropertyValue(cmbCondPropPercent,"Text",false).ToString().Trim());
                    SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                    this.SetLabelValue(m_frmTherm.lblMsg, "Text", "Calculate Adjustment Factors For RsCd=" + m_strCurrFIADBRsCd + " and EvalId=" + m_strCurrFIADBEvalId);
                    for (x = 0; x <= strSql.Length - 1; x++)
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, strSql[x] +  "\r\n");
                        m_ado.SqlNonQuery(oConn, strSql[x]);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 20 + x + 5);
                        if (m_ado.m_intError != 0) break;
                    }
                    m_ado.CloseConnection(oConn);    

                    m_intError = m_ado.m_intError;
                    SetThermValue(m_frmTherm.progressBar1, "Value", 100);
                    System.Threading.Thread.Sleep(2000);
                    SetThermValue(m_frmTherm.progressBar2, "Value", 60);
                }
                
                //create tablelinks to the projects main folder
                if (this.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                    this.SetLabelValue(m_frmTherm.lblMsg, "Text", "Create table links");
                    //create a table link from the newly created BIOSUM_ADJFACTORS into the master.mdb

                    //instatiate dao for creating links in the temp table
                    //to the fiadb plot, cond, and tree input tables
                    dao_data_access oDao = new dao_data_access();
                    SetThermValue(m_frmTherm.progressBar1, "Value", 10);
                    //create links to the fiadb input tables in the temp mdb file
                    //plot table
                    strFIADBDbFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBFiadbInputFile, "Text", false);
                    strFIADBDbFile = strFIADBDbFile.Trim();
                    strSourceTableName = "BIOSUM_PLOT";
                    strDestTableLinkName = "fiadb_plot_input";
                    oDao.CreateTableLink(this.m_strTempMDBFile,strDestTableLinkName,strFIADBDbFile,strSourceTableName,true);
                    SetThermValue(m_frmTherm.progressBar1, "Value", 40);
                    //cond table
                    strSourceTableName = "BIOSUM_COND";
                    strDestTableLinkName = "fiadb_cond_input";
                    if (oDao.m_intError==0) oDao.CreateTableLink(this.m_strTempMDBFile, strDestTableLinkName, strFIADBDbFile, strSourceTableName, true);
                    SetThermValue(m_frmTherm.progressBar1, "Value", 70);
                    //biosum adjustment factors table
                    strSourceTableName = frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName;
                    strDestTableLinkName = "fiadb_biosum_adjustment_factors_input";
                    if (oDao.m_intError==0) oDao.CreateTableLink(this.m_strTempMDBFile, strDestTableLinkName, strFIADBDbFile, strSourceTableName, true);




                    m_intError = oDao.m_intError;
                    
                    //destroy the object and release it from memory
                    oDao.m_DaoWorkspace.Close();
                    oDao = null;

                     m_intError = m_ado.m_intError;
                    SetThermValue(m_frmTherm.progressBar1, "Value", 100);
                    System.Threading.Thread.Sleep(2000);
                    SetThermValue(m_frmTherm.progressBar2, "Value", 70);

                  
                    
                }
                //delete any records from the production biosum adjustment factor table that did not previously complete processing (error or user cancelled)
                //or any previous rscd and evalid that equal the current ones
                if (this.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                    this.SetLabelValue(m_frmTherm.lblMsg, "Text", "Deleting Old Data");
                    //open the connection to the temp mdb file 
                    this.m_ado.OpenConnection(m_ado.getMDBConnString(this.m_strTempMDBFile,"",""), ref oConn);
                    m_ado.m_strSQL = "DELETE FROM " + this.m_strBiosumPopStratumAdjustmentFactorsTable + " WHERE biosum_status_cd=9 or (rscd=" + m_strCurrFIADBRsCd + " and evalid=" + m_strCurrFIADBEvalId + ")";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL+ "\r\n");
                    m_ado.SqlNonQuery(oConn, m_ado.m_strSQL);
                    SetThermValue(m_frmTherm.progressBar1, "Value", 50);
                    

                    m_intError = m_ado.m_intError;
                }
                //append work table to production table
                if (this.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    this.SetLabelValue(m_frmTherm.lblMsg, "Text", "Appending New Data");
                    //delete any previous rscd and evalid that equal the current ones
                    m_ado.m_strSQL = "INSERT INTO " + this.m_strBiosumPopStratumAdjustmentFactorsTable + " SELECT a.*,9 AS biosum_status_cd FROM fiadb_biosum_adjustment_factors_input a";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
                    m_ado.SqlNonQuery(oConn, m_ado.m_strSQL);
                    SetThermValue(m_frmTherm.progressBar1, "Value", 100);
                    SetThermValue(m_frmTherm.progressBar2, "Value", 100);
                     System.Threading.Thread.Sleep(2000);

                }


                if (this.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {

                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Visible", true);
                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);
                }

                CalculateAdjustments_Finish();

            }
            catch (System.Threading.ThreadInterruptedException err)
            {
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (System.Threading.ThreadAbortException err)
            {
                if (oConn != null)
                {
                    if (oConn.State != System.Data.ConnectionState.Closed)
                    {
                        m_ado.CloseConnection(oConn);
                    }
                    oConn = null;
                }
                if (m_ado != null)
                {
                    if (m_ado.m_DataSet != null)
                    {
                        this.m_ado.m_DataSet.Clear();
                        this.m_ado.m_DataSet.Dispose();
                    }
                    this.m_ado = null;
                }
                this.CancelThreadCleanup();
                this.ThreadCleanUp();
                this.CleanupThread();

            }
            catch (Exception err)
            {
                MessageBox.Show("!!Error!! \n" +
                    "Module - uc_plot_input.CalculateAdjustments_Process  \n" +
                    "Err Msg - " + err.Message.ToString().Trim(),
                    "FVS Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }
            finally
            {

            }

            if (oConn != null)
            {
                if (oConn.State != System.Data.ConnectionState.Closed)
                {
                    m_ado.CloseConnection(oConn);
                }
                oConn = null;
            }
            if (m_ado != null)
            {
                if (m_ado.m_DataSet != null)
                {
                    this.m_ado.m_DataSet.Clear();
                    this.m_ado.m_DataSet.Dispose();
                }
                this.m_ado = null;
            }
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);
            if (this.m_frmTherm != null) frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", false);




            CalculateAdjustments_Finish();

            CleanupThread();

            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
        }
		private void LoadMDBPlotCondTreeData_Process()
        {
            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
			string strBiosumPlotId="";
			string strFields="";
			
			int x=0;
			int y=0;
			string strCol="";
            string str="";
            string str2 = "";

			string strSourceTableLink="";
			//string strMsg="";
            string strFIADBDbFile = "";
            string strSourceTableName = "";
            string strDestTableLinkName = "";

            m_intAddedPlotRows=0;
		    m_intAddedCondRows=0;
		    m_intAddedTreeRows=0;
			m_intAddedSiteTreeRows=0;

		    this.m_intError=0;		
			

			//-----------PREPARATION FOR ADDING PLOT RECORDS---------//

            try
            {
                //instatiate the oledb data access class
                this.m_ado = new ado_data_access();

                //progress bar 1: single process
                this.SetThermValue(m_frmTherm.progressBar1,"Maximum", 100);
                this.SetThermValue(m_frmTherm.progressBar1,"Minimum", 0);
                this.SetThermValue(m_frmTherm.progressBar1,"Value", 0);
                this.SetLabelValue(m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);

                //progress bar 2: overall progress
                this.SetThermValue(m_frmTherm.progressBar2, "Maximum", 100);
                this.SetThermValue(m_frmTherm.progressBar2, "Minimum", 0);
                this.SetThermValue(m_frmTherm.progressBar2, "Value", 0);
                this.SetLabelValue(m_frmTherm.lblMsg2, "Text", "Overall Progress");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);




                //create a temporary mdb file with links to all the project tables
                //and return the name of the file that contains the links
                this.m_strTempMDBFile = m_oDatasource.CreateMDBAndTableDataSourceLinks();


                //instatiate dao for creating links in the temp table
                //to the fiadb plot, cond, and tree input tables
                dao_data_access p_dao1 = new dao_data_access();
                this.SetLabelValue(m_frmTherm.lblMsg, "Text", "Creating Datasource Links");
                //create links to the fiadb input tables in the temp mdb file
                //plot table
                //str = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBFiadbInputFile, "Text", false);
                //str2 = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.ComboBox)cmbFiadbPlotTable, "Text", false);
                //p_dao1.CreateTableLink(this.m_strTempMDBFile, "fiadb_plot_input", str.Trim(), str2.Trim());
                //cond table
                //str2 = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.ComboBox)cmbFiadbCondTable, "Text", false);
                //p_dao1.CreateTableLink(this.m_strTempMDBFile, "fiadb_cond_input", str.Trim(), str2.Trim());
                //plot table
                strFIADBDbFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBFiadbInputFile, "Text", false);
                strFIADBDbFile = strFIADBDbFile.Trim();
                strSourceTableName = "BIOSUM_PLOT";
                strDestTableLinkName = "fiadb_plot_input";
                p_dao1.CreateTableLink(this.m_strTempMDBFile, strDestTableLinkName, strFIADBDbFile, strSourceTableName, true);
                
                //cond table
                strSourceTableName = "BIOSUM_COND";
                strDestTableLinkName = "fiadb_cond_input";
                if (p_dao1.m_intError == 0) p_dao1.CreateTableLink(this.m_strTempMDBFile, strDestTableLinkName, strFIADBDbFile, strSourceTableName, true);
                //tree table
                str2 = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.ComboBox)cmbFiadbTreeTable, "Text", false);
                if (p_dao1.m_intError == 0) p_dao1.CreateTableLink(this.m_strTempMDBFile, "fiadb_tree_input", strFIADBDbFile, str2.Trim());
                //tree regional biomass
                str2 = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.ComboBox)cmbFiadbTreeRegionalBiomassTable, "Text", false);
                if (p_dao1.m_intError == 0) p_dao1.CreateTableLink(this.m_strTempMDBFile, "fiadb_treeRegionalBiomass_input", strFIADBDbFile, str2.Trim());
                //site tree
                str2 = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.ComboBox)cmbFiadbSiteTreeTable, "Text", false);
                if (p_dao1.m_intError == 0) p_dao1.CreateTableLink(this.m_strTempMDBFile, "fiadb_site_tree_input", strFIADBDbFile, str2.Trim());
                //biosum_volume table
                //ORACLE FCS Tree Volume Table
                //create a temporary link to the ORACLE FCS BIOSUM_VOLUME table
                if (p_dao1.m_intError == 0)
                {
                    System.Threading.Thread.Sleep(1000);
                    if (p_dao1.TableExists(m_strTempMDBFile, "fcs_biosum_volume"))
                    {
                        p_dao1.DeleteTableFromMDB(m_strTempMDBFile, "fcs_biosum_volume");
                    }

                    for (int z = 1; z <= 5; z++)
                    {
                        System.Threading.Thread.Sleep(2000 * z);
                        p_dao1.m_intError = 0;
                        p_dao1.CreateOracleXETableLink("FIA Biosum Oracle Services", "FCS", "fcs", "FCS", "BIOSUM_VOLUME", m_strTempMDBFile.Trim(), "fcs_biosum_volume");
                        if (p_dao1.m_intError == 0) break;
                    }
                    if (p_dao1.m_intError != 0)
                    {
                        MessageBox.Show("!!Failed to create Oracle XE ODBC table link!! Contact technical support", "FIA Biosum");
                    }
                }


                m_intError = p_dao1.m_intError;

                //destroy the object and release it from memory
                p_dao1.m_DaoWorkspace.Close();
                p_dao1 = null;


                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 10);

                System.Data.DataTable dtPlotSchema = new DataTable();
                System.Data.DataTable dtCondSchema = new DataTable();
                System.Data.DataTable dtTreeSchema = new DataTable();
                System.Data.DataTable dtSiteTreeSchema = new DataTable();
                System.Data.DataTable dtFIADBPlotSchema = new DataTable();
                System.Data.DataTable dtFIADBCondSchema = new DataTable();
                System.Data.DataTable dtFIADBTreeSchema = new DataTable();
                System.Data.DataTable dtFIADBSiteTreeSchema = new DataTable();
               

                    //get an ado connection string for the temp mdb file
                    this.m_strTempMDBFileConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile, "", "");


                    //create a new connection to the temp MDB file
                    this.m_connTempMDBFile = new System.Data.OleDb.OleDbConnection();

                    //open the connection to the temp mdb file 
                    this.m_ado.OpenConnection(this.m_strTempMDBFileConn, ref this.m_connTempMDBFile);



                    m_ado.m_strSQL = "DELETE FROM " + this.m_strPlotTable + " WHERE biosum_status_cd=9 OR LEN(biosum_plot_id)=0;";
                    m_ado.SqlNonQuery(m_connTempMDBFile, m_ado.m_strSQL);
                    m_ado.m_strSQL = "DELETE FROM " + this.m_strCondTable + " WHERE biosum_status_cd=9;";
                    m_ado.SqlNonQuery(m_connTempMDBFile, m_ado.m_strSQL);
                    m_ado.m_strSQL = "DELETE FROM " + this.m_strTreeTable + " WHERE biosum_status_cd=9;";
                    m_ado.SqlNonQuery(m_connTempMDBFile, m_ado.m_strSQL);
                    m_ado.m_strSQL = "DELETE FROM " + this.m_strTreeRegionalBiomassTable + " WHERE biosum_status_cd=9;";
                    m_ado.SqlNonQuery(m_connTempMDBFile, m_ado.m_strSQL);
                    m_ado.m_strSQL = "DELETE FROM " + this.m_strSiteTreeTable + " WHERE biosum_status_cd=9;";
                    m_ado.SqlNonQuery(m_connTempMDBFile, m_ado.m_strSQL);

                    if (m_intError == 0) m_intError = m_ado.m_intError;



                if (this.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm,"AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 20);

                    /****************************************************************
                     **get the table structure that results from executing the sql
                     ****************************************************************/
                    //get the fiabiosum table structures
                    dtPlotSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strPlotTable);
                    dtCondSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strCondTable);
                    dtTreeSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strTreeTable);
                    dtSiteTreeSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strSiteTreeTable);
                    //get the fiadb table structures
                    dtFIADBPlotSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from fiadb_plot_input");
                    dtFIADBCondSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from fiadb_cond_input");
                    dtFIADBTreeSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from fiadb_tree_input");
                    dtFIADBSiteTreeSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from fiadb_site_tree_input");
                    m_intError = m_ado.m_intError;
                }

                if (this.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm,"AbortProcess"))
                {
                   


                    //-------------PLOT TABLE----------------//
                    strSourceTableLink = "fiadb_plot_input";

                    //build field list string to insert sql by matching 
                    //up the column names in the biosum plot table and the fiadb plot table
                    strFields = "";
                    for (x = 0; x <= dtPlotSchema.Rows.Count - 1; x++)
                    {
                        strCol = dtPlotSchema.Rows[x]["columnname"].ToString().Trim();
                        //see if there is an equivalent FIADB column
                        for (y = 0; y <= dtFIADBPlotSchema.Rows.Count - 1; y++)
                        {
                            if (strCol.Trim().ToUpper() == dtFIADBPlotSchema.Rows[y]["columnname"].ToString().ToUpper())
                            {
                                if (strFields.Trim().Length == 0)
                                {
                                    strFields = strCol;
                                }
                                else
                                {
                                    strFields += "," + strCol;
                                }
                                break;
                            }
                        }
                    }
                    SetLabelValue(m_frmTherm.lblMsg,"Text","Plot Table: Insert New Plot Records");
                    if (Checked(rdoFilterByFile) == true && m_strPlotIdList.Trim().Length > 0 &&
                        !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        string strDelimiter = ",";
                        string[] strPlotIdArray = m_strPlotIdList.Split(strDelimiter.ToCharArray());
                        this.m_ado.m_strSQL = "SELECT CN INTO input_cn FROM " + this.m_strPlotTable + " WHERE 1=2";
                        this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                        for (x = 0; x <= strPlotIdArray.Length - 1; x++)
                        {
                            if (strPlotIdArray[x] != null && strPlotIdArray[x].Trim().Length > 0)
                            {
                                this.m_ado.m_strSQL = "INSERT INTO input_cn (CN) VALUES (" + strPlotIdArray[x].Trim() + ")";
                                this.m_ado.SqlNonQuery(this.m_connTempMDBFile, m_ado.m_strSQL);
                            }
                        }
                        this.m_ado.m_strSQL = "SELECT '999999999999999999999999' AS biosum_plot_id,9 AS biosum_status_cd,p.* INTO tempplot " +
                            " FROM " + strSourceTableLink + " p," +
                            this.m_strPpsaTable + " ppsa, " +
                            "input_cn " +
                            " WHERE p.cn=input_cn.cn AND " +
                            "p.cn=ppsa.plt_cn AND " +
                            "ppsa.rscd = " + this.m_strCurrFIADBRsCd + " AND " +
                            "ppsa.evalid = " + this.m_strCurrFIADBEvalId;

                    }
                    else
                    {

                        /********************************************************
                         **create plot input insert command
                         ********************************************************/
                        //check the user defined filters
                        this.m_ado.m_strSQL = "SELECT '999999999999999999999999' AS biosum_plot_id,9 AS biosum_status_cd,p.* INTO tempplot FROM " + strSourceTableLink + " p " +
                            " INNER JOIN " + this.m_strPpsaTable + " ppsa ON p.cn=ppsa.plt_cn " +
                            " WHERE ppsa.rscd = " + this.m_strCurrFIADBRsCd + " AND " +
                            "ppsa.evalid = " + this.m_strCurrFIADBEvalId;
                    }
                    if (Checked(rdoFilterNone))
                    {
                        //forested/nonforested filters
                        if (Checked(chkNonForested) &&
                            Checked(chkForested))
                        {
                            //all plots

                        }
                        else if (Checked(chkForested))
                        {
                            //forested plots
                            this.m_ado.m_strSQL = m_ado.m_strSQL + " AND p.plot_status_cd = 1";
                        }
                        else
                        {
                            //nonforested plots
                            this.m_ado.m_strSQL = m_ado.m_strSQL + " AND p.plot_status_cd IS NULL OR p.plot_status_cd <> 1";
                        }
                    }
                   
                    else if (Checked(rdoFilterByMenu))
                    {
                        if (Checked(chkNonForested) &&
                            Checked(chkForested))
                        {
                            //all plots
                        }
                        else if (Checked(chkForested))
                        {
                            //forested plots
                            this.m_ado.m_strSQL = m_ado.m_strSQL + " AND (p.plot_status_cd = 1) ";
                        }
                        else
                        {
                            //nonforested plots
                            this.m_ado.m_strSQL = m_ado.m_strSQL + " AND (p.plot_status_cd IS NULL OR p.plot_status_cd <> 1) ";
                        }
                        if (this.m_strStateCountyPlotSQL.Trim().Length > 0)
                        {
                            //state,county,plot filter
                            this.BuildFilterByPlotString("ppsa.statecd", "ppsa.countycd", "ppsa.plot", false);
                            this.m_ado.m_strSQL += " AND " + this.m_strStateCountyPlotSQL.Trim() + ";";
                        }
                        else
                        {
                            //state,county filter
                            this.BuildFilterByStateCountyString("ppsa.statecd", "ppsa.countycd", false);
                            this.m_ado.m_strSQL += " AND " + this.m_strStateCountySQL.Trim() + ";";
                        }

                    }

                    //insert new plot records
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    m_intError = m_ado.m_intError;
                }
                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 30);
                    SetLabelValue(m_frmTherm.lblMsg,"Text","Plot Table: Update Biosum_Plot_Id Column");
                    
                    System.Data.OleDb.OleDbTransaction p_transTempPlot = this.m_connTempMDBFile.BeginTransaction();

                    //update the biosum_plot_id column
                    this.m_ado.m_DataSet = new DataSet("FIADB");
                    this.m_ado.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
                    this.m_ado.AddSQLQueryToDataSet(this.m_connTempMDBFile,
                        ref m_ado.m_OleDbDataAdapter,
                        ref m_ado.m_DataSet,
                        ref p_transTempPlot,
                        "SELECT * FROM tempplot", "tempplot");

                    this.m_ado.ConfigureDataAdapterUpdateCommand(this.m_connTempMDBFile,
                        m_ado.m_OleDbDataAdapter,
                        p_transTempPlot,
                        "SELECT biosum_plot_id FROM tempplot",
                        "SELECT CN FROM tempplot",
                        "tempplot");

                    for (x = 0; x <= this.m_ado.m_DataSet.Tables["tempplot"].Rows.Count - 1; x++)
                    {
                        strBiosumPlotId = this.CreateBiosumPlotId(this.m_ado.m_DataSet.Tables["tempplot"].Rows[x]);
                        this.m_ado.m_DataSet.Tables["tempplot"].Rows[x].BeginEdit();
                        this.m_ado.m_DataSet.Tables["tempplot"].Rows[x]["biosum_plot_id"] = strBiosumPlotId;
                        this.m_ado.m_DataSet.Tables["tempplot"].Rows[x].EndEdit();
                    }
                    m_ado.m_OleDbDataAdapter.Update(this.m_ado.m_DataSet.Tables["tempplot"]);
                    p_transTempPlot.Commit();
                    this.m_ado.m_DataSet.Tables["tempplot"].AcceptChanges();
                    p_transTempPlot = null;
                    m_ado.m_OleDbDataAdapter.Dispose();
                    m_ado.m_OleDbDataAdapter = null;
                    m_intError = m_ado.m_intError;
                }
                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 40);
                    
                    //insert the new plot records into the plot table
                    m_ado.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (biosum_plot_id,biosum_status_cd," + strFields + ") " +
                        "SELECT TRIM(biosum_plot_id),biosum_status_cd," + strFields + " FROM tempplot";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    m_intError = m_ado.m_intError;
                }
                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {

                    //initialize columns
                    m_ado.m_strSQL = "UPDATE " + this.m_strPlotTable + " " +
                        "SET gis_protected_area_yn='N'," +
                        "gis_roadless_yn='N'," +
                        "all_cond_not_accessible_yn='N'," +
                        "plot_accessible_yn='Y'," +
                        "gis_status_id=1 " +
                        "WHERE biosum_status_cd=9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    m_intError = m_ado.m_intError;
                }
                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {

                    //create plot column update work table
                    this.m_strSQL = "SELECT biosum_plot_id, statecd as cond_ttl " +
                        "INTO plot_column_updates_work_table FROM " + this.m_strPlotTable.Trim() + " WHERE 1=2;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    System.Threading.Thread.Sleep(10000);
                    m_intError = m_ado.m_intError;
                }

                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1,"Value",40);
                    SetThermValue(m_frmTherm.progressBar2,"Value",20);

                    
                    //-------------CONDITION TABLE----------------//
                    strSourceTableLink = "fiadb_cond_input";



                    //build field list string to insert sql by matching 
                    //up the column names in the biosum plot table and the fiadb plot table
                    strFields = "";
                    for (x = 0; x <= dtCondSchema.Rows.Count - 1; x++)
                    {
                        strCol = dtCondSchema.Rows[x]["columnname"].ToString().Trim();
                        //see if there is an equivalent FIADB column
                        for (y = 0; y <= dtFIADBCondSchema.Rows.Count - 1; y++)
                        {
                            if (strCol.Trim().ToUpper() == dtFIADBCondSchema.Rows[y]["columnname"].ToString().ToUpper())
                            {
                                if (strFields.Trim().Length == 0)
                                {
                                    strFields = strCol;
                                }
                                else
                                {
                                    strFields += "," + strCol;
                                }
                                break;
                            }
                        }
                    }

                    /********************************************************
                     **create condition input insert command
                     ********************************************************/
                    //check the user defined filters
                    SetLabelValue(m_frmTherm.lblMsg,"Text","Condition Table: Insert New  Records");
                    this.m_ado.m_strSQL = "SELECT p.biosum_plot_id, TRIM(p.biosum_plot_id) + TRIM(CSTR(c.condid)) AS biosum_cond_id,9 AS biosum_status_cd,c.* INTO tempcond FROM " + strSourceTableLink + " c " +
                        " INNER JOIN " + this.m_strPlotTable + " p ON c.plt_cn=p.cn WHERE " +
                        " p.biosum_status_cd=9";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    m_intError = m_ado.m_intError;
                }
                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 60);

                    //insert the new condition records into the condition table
                    m_ado.m_strSQL = "INSERT INTO " + this.m_strCondTable + " (biosum_plot_id,biosum_cond_id,biosum_status_cd," + strFields + ") " +
                        "SELECT TRIM(biosum_plot_id),TRIM(biosum_cond_id),biosum_status_cd," + strFields + " FROM tempcond";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);

                    m_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " d " +
                        "INNER JOIN " + strSourceTableLink + " s " +
                        "ON d.cn = s.cn " +
                        "SET d.landclcd = s.cond_status_cd";
                    if (m_ado.m_intError == 0)
                        this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);


                    //create cond column work table
                    this.m_strSQL = "SELECT biosum_cond_id, qmd_tot_cm,hwd_qmd_tot_cm," +
                        "swd_qmd_tot_cm,tpacurr,hwd_tpacurr,swd_tpacurr,ba_ft2_ac," +
                        "hwd_ba_ft2_ac,swd_ba_ft2_ac,vol_ac_grs_stem_ttl_ft3," +
                        "hwd_vol_ac_grs_stem_ttl_ft3,swd_vol_ac_grs_stem_ttl_ft3," +
                        "vol_ac_grs_ft3, hwd_vol_ac_grs_ft3," +
                        "swd_vol_ac_grs_ft3,volcsgrs," +
                        "hwd_volcsgrs, swd_volcsgrs INTO cond_column_updates_work_table " +
                        "FROM " + this.m_strCondTable.Trim() + " WHERE 1=2;";
                    if (m_ado.m_intError == 0)
                        m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    m_intError = m_ado.m_intError;
                }
                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 65);
                    SetThermValue(m_frmTherm.progressBar2, "Value", 40);
                    



                    //-------------TREE TABLE----------------//
                    strSourceTableLink = "fiadb_tree_input";


                    //build field list string to insert sql by matching 
                    //up the column names in the biosum plot table and the fiadb plot table
                    strFields = "";
                    for (x = 0; x <= dtTreeSchema.Rows.Count - 1; x++)
                    {
                        strCol = dtTreeSchema.Rows[x]["columnname"].ToString().Trim();
                        //see if there is an equivalent FIADB column
                        for (y = 0; y <= dtFIADBTreeSchema.Rows.Count - 1; y++)
                        {
                            if (strCol.Trim().ToUpper() == dtFIADBTreeSchema.Rows[y]["columnname"].ToString().ToUpper())
                            {
                                if (strFields.Trim().Length == 0)
                                {
                                    strFields = strCol;
                                }
                                else
                                {
                                    strFields += "," + strCol;
                                }
                                break;
                            }
                        }
                    }

                    /********************************************************
                     **create tree input insert command
                     ********************************************************/
                    //check the user defined filters
                    SetLabelValue(m_frmTherm.lblMsg,"Text","Tree Table: Insert New  Records");

                    this.m_ado.m_strSQL = "SELECT TRIM(p.biosum_plot_id) + TRIM(CSTR(t.condid)) AS biosum_cond_id,9 AS biosum_status_cd,t.* INTO temptree FROM " + strSourceTableLink + " t " +
                        " INNER JOIN " + this.m_strPlotTable + " p ON t.plt_cn=p.cn " +
                        " WHERE p.biosum_status_cd=9";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    m_intError = m_ado.m_intError;
                }
                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 75);
                    //insert the new condition records into the condition table
                    m_ado.m_strSQL = "INSERT INTO " + this.m_strTreeTable + " (biosum_cond_id,biosum_status_cd," + strFields + ") " +
                        "SELECT TRIM(biosum_cond_id),biosum_status_cd," + strFields + " FROM temptree";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);

                    m_ado.m_strSQL = "INSERT INTO " + this.m_strTreeRegionalBiomassTable + "  " +
                        "SELECT s.tre_cn,s.statecd," +
                        "s.regional_drybiot,s.regional_drybiom," +
                        "9 AS biosum_status_cd FROM fiadb_treeRegionalBiomass_input s " +
                        "INNER JOIN " + this.m_strTreeTable + " t " +
                        "ON t.cn = s.tre_cn WHERE t.biosum_status_cd=9";
                    if (m_ado.m_intError == 0)
                        this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    m_intError = m_ado.m_intError;
                }
                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 80);

                    //update the cullbf column
                    this.m_ado.m_strSQL = "UPDATE " + this.m_strTreeTable + " " +
                        "SET cullbf=IIF(cullbf IS NULL," +
                        "IIF(cull IS NOT NULL AND roughcull IS NOT NULL," +
                        "cull + roughcull," +
                        "IIF(cull IS NOT NULL,cull," +
                        "IIF(roughcull IS NOT NULL,roughcull,0))),cullbf)";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    m_intError = m_ado.m_intError;
                }
                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {

                    //-------------SITE TREE TABLE----------------//
                    strSourceTableLink = "fiadb_site_tree_input";


                    //build field list string to insert sql by matching 
                    //up the column names in the biosum plot table and the fiadb plot table
                    strFields = "";
                    for (x = 0; x <= dtSiteTreeSchema.Rows.Count - 1; x++)
                    {
                        strCol = dtSiteTreeSchema.Rows[x]["columnname"].ToString().Trim();
                        //see if there is an equivalent FIADB column
                        for (y = 0; y <= dtFIADBSiteTreeSchema.Rows.Count - 1; y++)
                        {
                            if (strCol.Trim().ToUpper() == dtFIADBSiteTreeSchema.Rows[y]["columnname"].ToString().ToUpper())
                            {
                                if (strFields.Trim().Length == 0)
                                {
                                    strFields = strCol;
                                }
                                else
                                {
                                    strFields += "," + strCol;
                                }
                                break;
                            }
                        }
                    }

                    /********************************************************
                     **create site tree input insert command
                     ********************************************************/
                    //check the user defined filters
                    SetLabelValue(m_frmTherm.lblMsg,"Text","Site Tree Table: Insert New  Records");
                    this.m_ado.m_strSQL = "SELECT TRIM(p.biosum_plot_id) AS biosum_plot_id,9 AS biosum_status_cd,t.* INTO tempsitetree FROM " + strSourceTableLink + " t " +
                        " INNER JOIN " + this.m_strPlotTable + " p ON t.plt_cn=p.cn " +
                        " WHERE p.biosum_status_cd=9";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    m_intError = m_ado.m_intError;
                }
                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 80);
                    

                    //insert the new condition records into the condition table
                    m_ado.m_strSQL = "INSERT INTO " + this.m_strSiteTreeTable + " (biosum_plot_id,biosum_status_cd," + strFields + ") " +
                        "SELECT TRIM(biosum_plot_id),biosum_status_cd," + strFields + " FROM tempsitetree";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);

                    m_intError = m_ado.m_intError;
                }

                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 100);
                    SetThermValue(m_frmTherm.progressBar2, "Value", 80);
                    this.UpdateColumns(m_ado);
                    m_intError = m_ado.m_intError;
                }
                if (this.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    m_intAddedPlotRows = (int)this.m_ado.getRecordCount(this.m_connTempMDBFile, "select count(*) from " + this.m_strPlotTable + " WHERE biosum_status_cd=9", m_strPlotTable);
                    m_intAddedCondRows = (int)this.m_ado.getRecordCount(this.m_connTempMDBFile, "select count(*) from " + this.m_strCondTable + " WHERE biosum_status_cd=9", m_strCondTable);
                    m_intAddedTreeRows = (int)this.m_ado.getRecordCount(this.m_connTempMDBFile, "select count(*) from " + this.m_strTreeTable + " WHERE biosum_status_cd=9", m_strTreeTable);
                    m_intAddedTreeRegionalBiomassRows = (int)this.m_ado.getRecordCount(this.m_connTempMDBFile, "select count(*) from " + this.m_strTreeRegionalBiomassTable + " WHERE biosum_status_cd=9", m_strTreeRegionalBiomassTable);
                    m_intAddedSiteTreeRows = (int)this.m_ado.getRecordCount(this.m_connTempMDBFile, "select count(*) from " + this.m_strSiteTreeTable + " WHERE biosum_status_cd=9", m_strSiteTreeTable);

                    this.m_strSQL = " UPDATE " + this.m_strPlotTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    this.m_strSQL = " UPDATE " + this.m_strCondTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    this.m_strSQL = " UPDATE " + this.m_strTreeTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    this.m_strSQL = " UPDATE " + this.m_strPopEvalTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    this.m_strSQL = " UPDATE " + this.m_strTreeRegionalBiomassTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    this.m_strSQL = " UPDATE " + this.m_strPopStratumTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    this.m_strSQL = " UPDATE " + this.m_strPpsaTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    this.m_strSQL = " UPDATE " + this.m_strPopEstUnitTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    this.m_strSQL = " UPDATE " + this.m_strSiteTreeTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    this.m_strSQL = " UPDATE " + this.m_strBiosumPopStratumAdjustmentFactorsTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);

                    SetThermValue(m_frmTherm.progressBar1, "Value", GetThermValue(m_frmTherm.progressBar1, "Maximum"));
                    SetThermValue(m_frmTherm.progressBar2, "Value", GetThermValue(m_frmTherm.progressBar2, "Maximum"));
                    frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Button)m_frmTherm.btnCancel, "Visible", false);


                    MessageBox.Show("Successfully Appended \n" +
                        m_intAddedPlotRows.ToString().Trim() + " Plot Records \n" +
                        m_intAddedCondRows.ToString().Trim() + " Condition Records \n" +
                        m_intAddedTreeRows.ToString().Trim() + " Tree Records \n" +
                        m_intAddedTreeRegionalBiomassRows.ToString().Trim() + " Tree Regional Biomass Records \n" +
                        m_intAddedSiteTreeRows.ToString().Trim() + " Site Tree Records", "Add Plot Data");

                    this.m_strLoadedPopEstUnitInputTable =
                        (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.ComboBox)cmbFiadbPopEstUnitTable, "Text", false);
                    this.m_strLoadedPopStratumInputTable =
                        (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.ComboBox)cmbFiadbPopStratumTable, "Text", false);
                    this.m_strLoadedPpsaInputTable =
                        (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.ComboBox)cmbFiadbPpsaTable, "Text", false);
                    this.m_strLoadedFIADBEvalId = this.m_strCurrFIADBEvalId;
                    this.m_strLoadedFIADBRsCd = this.m_strCurrFIADBRsCd;
                    this.m_strLoadedFiadbInputFile =
                        (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBFiadbInputFile, "Text", false);
                    System.Threading.Thread.Sleep(1000);

                }
                else
                {
                    //error occurred in the updatecolumns so delete the records
                    this.m_strSQL = "DELETE FROM " + this.m_strPlotTable + " WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    //delete added cond records since error occured
                    this.m_strSQL = "DELETE FROM " + this.m_strCondTable + " WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    //delete added cond records since error occured
                    this.m_strSQL = "DELETE FROM " + this.m_strTreeTable + " WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    //delete added cond records since error occured
                    this.m_strSQL = "DELETE FROM " + this.m_strTreeRegionalBiomassTable + " WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    this.m_strSQL = "DELETE FROM " + this.m_strPopEvalTable + " WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    this.m_strSQL = "DELETE FROM " + this.m_strPopStratumTable + " WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    this.m_strSQL = "DELETE FROM " + this.m_strPpsaTable + " WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    this.m_strSQL = "DELETE FROM " + this.m_strPopEstUnitTable + " WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    this.m_strSQL = "DELETE FROM " + this.m_strSiteTreeTable + " WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    this.m_strSQL = "DELETE FROM " + this.m_strBiosumPopStratumAdjustmentFactorsTable + " WHERE biosum_status_cd = 9;";
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    MessageBox.Show("!!Error Occured Adding Plot Records: 0 Records Added!!", "Add Plot Data", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                }


                this.m_connTempMDBFile.Close();
                while (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
                    System.Threading.Thread.Sleep(1000);
                this.m_ado.m_DataSet.Clear();
                this.m_ado.m_DataSet.Dispose();
                this.m_ado = null;

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Visible", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);


                LoadMDBPlotCondTreeData_Finish();

            }
            catch (System.Threading.ThreadInterruptedException err)
			{
				MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
			}
			catch  (System.Threading.ThreadAbortException err)
			{
                if (this.m_connTempMDBFile != null)
                {
                    if (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
                    {
                        m_ado.CloseConnection(m_connTempMDBFile);
                    }
                    m_connTempMDBFile = null;
                }
                if (m_ado != null)
                {
                    if (m_ado.m_DataSet != null)
                    {
                        this.m_ado.m_DataSet.Clear();
                        this.m_ado.m_DataSet.Dispose();
                    }
                    this.m_ado = null;
                }
                this.CancelThreadCleanup();
			    this.ThreadCleanUp();
				this.CleanupThread();
				
			}
			catch (Exception err)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_plot_input:mdbFIADBFileInput  \n" + 
					"Err Msg - " + err.Message.ToString().Trim(),
					"FVS Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}
			finally
			{
               
			}
			
			if (this.m_connTempMDBFile != null)
			{
                if (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
                {
                    m_ado.CloseConnection(m_connTempMDBFile);
                }
                m_connTempMDBFile = null;
			}
			if (m_ado != null)
			{
				if (m_ado.m_DataSet != null)
				{
					this.m_ado.m_DataSet.Clear();
					this.m_ado.m_DataSet.Dispose();
				}
				this.m_ado = null;
			}
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);
			if (this.m_frmTherm != null) frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm,"Visible",false);

            

           
            LoadMDBPlotCondTreeData_Finish();

            CleanupThread();

            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
		
		
		

		


		}

		


		private void UpdateColumns(FIA_Biosum_Manager.ado_data_access p_ado)
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.UpdateColumns\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }

			//create work tables


            string strColumns = "";
            string strValues = "";
			string strTime = System.DateTime.Now.ToString();
				
			//----------------------COND COLUMN UPDATES-----------------------//
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//----------------------COND AND TREE COLUMN UPDATES-----------------------//\r\n");
			if (Checked(rdoFIADB))
			{
                SetThermValue(m_frmTherm.progressBar1,"Maximum",41);
                SetThermValue(m_frmTherm.progressBar1, "Minimum", 0);
                SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                


				if (Checked(rdoFIADB))
				{
                    SetLabelValue(m_frmTherm.lblMsg,"Text","Updating Condition Proportion Column...Stand By");
					frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
					//update the condition proportion column
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
									 "INNER JOIN ((" + this.m_strPpsaTable + " ppsa " + 
												 "INNER JOIN " + this.m_strPlotTable + " p " + 
												 "ON ppsa.plt_cn= p.cn) " + 
												 "INNER JOIN " + this.m_strBiosumPopStratumAdjustmentFactorsTable + " ps " + 
												 "ON ppsa.stratum_cn=ps.stratum_cn) " + 
									 "ON c.biosum_plot_id = p.biosum_plot_id " + 
									 "SET condprop = IIf(ps.pmh_macr Is Not Null And ps.pmh_macr>0," + 
														 "c.condprop_unadj/ps.pmh_macr," + 
													"IIf(ps.pmh_sub Is Not Null And ps.pmh_sub>0," + 
														 "c.condprop_unadj/ps.pmh_sub," + 
													"IIf(ps.pmh_micr Is Not Null And ps.pmh_micr>0," + 
														 "c.condprop_unadj/ps.pmh_micr,0)))";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

                    SetLabelValue(m_frmTherm.lblMsg,"Text","Updating Condition Acres Column...Stand By");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

					//update acres column
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
									 "INNER JOIN ((" + this.m_strPpsaTable + " ppsa " + 
												 "INNER JOIN " + this.m_strPlotTable + " p " + 
												 "ON ppsa.plt_cn= p.cn) " +
                                                 "INNER JOIN " + this.m_strBiosumPopStratumAdjustmentFactorsTable + " ps " + 
												 "ON ppsa.stratum_cn=ps.stratum_cn) " + 
									 "ON c.biosum_plot_id = p.biosum_plot_id " + 
									 "SET acres = IIF( c.condprop IS NOT NULL and ps.expns IS NOT NULL," + 
													"c.condprop * ps.expns,0)";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

                    //update condprop_specific column for when plot.macro_breakpoint_dia has a value
                    SetLabelValue(m_frmTherm.lblMsg, "Text", "Updating Tree Condprop Specific Column...Stand By");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
                    
                    p_ado.m_strSQL = "UPDATE " + m_strTreeTable + " t " +
                                        "INNER JOIN (" + m_strCondTable + " c INNER JOIN " + m_strPlotTable + " p ON c.biosum_plot_id=p.biosum_plot_id) " +
                                        "ON t.biosum_cond_id = c.biosum_cond_id " +
                                        "SET t.condprop_specific = " +
                                        "IIF(c.micrprop_unadj IS NOT NULL AND t.dia < 5," +
                                            "c.micrprop_unadj," +
                                        "IIF(c.subpprop_unadj IS NOT NULL AND " +
                                            "p.MACRO_BREAKPOINT_DIA IS NOT NULL AND " +
                                            "t.dia >= 5 AND " +
                                            "t.dia < p.MACRO_BREAKPOINT_DIA," +
                                            "c.subpprop_unadj," +
                                        "IIF(c.macrprop_unadj IS NOT NULL AND " +
                                            "p.MACRO_BREAKPOINT_DIA IS NOT NULL AND " +
                                            "t.dia >= p.MACRO_BREAKPOINT_DIA," +
                                            "c.macrprop_unadj))) " + 
                                        "WHERE t.biosum_status_cd=9";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
                    p_ado.SqlNonQuery(this.m_connTempMDBFile, p_ado.m_strSQL);
                    //check if null values 
                    p_ado.m_strSQL = "SELECT COUNT(*) AS ROWCOUNT " +
                                     "FROM " + m_strTreeTable + " t," +
                                               m_strTreeMacroPlotBreakPointDiaTable + " bp " +
                                     "WHERE t.biosum_status_cd=9 AND " +
                                           "t.condprop_specific IS NULL AND " +
                                           "t.statecd = bp.statecd AND " +
                                           "t.unitcd = bp.unitcd";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
                    //check if condprop_specific null and exists in the tree macro plot breakpoint diameter table
                    if ((double)p_ado.getSingleDoubleValueFromSQLQuery(m_connTempMDBFile, p_ado.m_strSQL, "temp") > 0)
                    {
                        //got some nulls
                        p_ado.m_strSQL = "UPDATE " + m_strTreeTable + " t " +
                                         "INNER JOIN ((" + m_strCondTable + " c " +
                                            "INNER JOIN " + m_strPlotTable + " p ON c.biosum_plot_id=p.biosum_plot_id) " +
                                            "INNER JOIN " + m_strTreeMacroPlotBreakPointDiaTable + " bp ON p.statecd = bp.statecd AND p.unitcd=bp.unitcd) " +
                                         "ON t.biosum_cond_id = c.biosum_cond_id " +
                                         "SET t.condprop_specific = " +
                                         "IIF(c.micrprop_unadj IS NOT NULL AND t.dia < 5," +
                                                "c.micrprop_unadj," +
                                         "IIF(c.subpprop_unadj IS NOT NULL AND " +
                                             "bp.MACRO_BREAKPOINT_DIA IS NOT NULL AND " +
                                             "t.dia >= 5 AND t.dia < bp.MACRO_BREAKPOINT_DIA," +
                                                "c.subpprop_unadj," +
                                         "IIF(c.macrprop_unadj IS NOT NULL AND " +
                                             "bp.MACRO_BREAKPOINT_DIA IS NOT NULL AND " +
                                             "t.dia >= bp.MACRO_BREAKPOINT_DIA," +
                                                "c.macrprop_unadj))) " +
                                         "WHERE t.biosum_status_cd=9";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
                        p_ado.SqlNonQuery(this.m_connTempMDBFile, p_ado.m_strSQL);
                    }
                    else
                    {
                        //see if we have nulls and no MACRO PLOT for the unit code
                         p_ado.m_strSQL = "SELECT COUNT(*) AS ROWCOUNT FROM tree a," + 
                                              "(SELECT t.* FROM tree t " + 
                                               "WHERE NOT EXISTS " + 
                                                   "(SELECT  * FROM TreeMacroPlotBreakPointDia bp " + 
                                                    "WHERE t.statecd=bp.statecd AND t.unitcd=bp.unitcd)) b " + 
                                         "WHERE a.CN=b.CN AND a.condprop_specific IS NULL AND a.biosum_status_cd=9";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
                        //handle for those states and units that do not have macro plot
                        if ((double)p_ado.getSingleDoubleValueFromSQLQuery(m_connTempMDBFile, p_ado.m_strSQL, "temp") > 0)
                        {
                            p_ado.m_strSQL = "UPDATE " + m_strTreeTable + " t " +
                                             "INNER JOIN (" + m_strCondTable + " c " +
                                             "INNER JOIN " + m_strPlotTable + " p ON c.biosum_plot_id=p.biosum_plot_id) " +
                                             "ON t.biosum_cond_id = c.biosum_cond_id AND t.condid = c.condid " +
                                             "SET t.condprop_specific = " +
                                             "IIF(c.micrprop_unadj IS NOT NULL AND t.dia < 5," +
                                                "c.micrprop_unadj," +
                                             "IIF(c.subpprop_unadj IS NOT NULL AND t.dia >= 5," +
                                                "c.subpprop_unadj)) " +
                                             "WHERE t.biosum_status_cd=9";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
                            p_ado.SqlNonQuery(this.m_connTempMDBFile, p_ado.m_strSQL);

                        }

                    }


                    SetLabelValue(m_frmTherm.lblMsg,"Text","Updating Tree tpacurr Column...Stand By");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
                    
                    

					//update tree tpacurr column
                    p_ado.m_strSQL = "UPDATE " + this.m_strTreeTable + " t " +
                        "SET tpacurr = IIF(t.tpa_unadj IS NOT NULL AND t.condprop_specific IS NOT NULL," +
                                       "t.tpa_unadj / t.condprop_specific,0)";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");

					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

                    SetLabelValue(m_frmTherm.lblMsg,"Text", "Updating Tree drybiom and drybiot Columns...Stand By");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

	
					//update tree drybiom and drybiot columns 
					p_ado.m_strSQL = "UPDATE " + this.m_strTreeTable + " t " + 
						"INNER JOIN " + this.m_strTreeRegionalBiomassTable + " drb " + 
						"ON t.cn = drb.tre_cn " + 
						"SET drybiom = IIF(drb.regional_drybiom IS NOT NULL,drb.regional_drybiom,null)," + 
						    "drybiot = IIF(drb.regional_drybiot IS NOT NULL,drb.regional_drybiot,null)";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);


				}
                SetLabelValue(m_frmTherm.lblMsg,"Text", "Updating Condition Table Columns...Stand By");
                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

				
                SetThermValue(m_frmTherm.progressBar1, "Value", 1);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					//tpa column
					//sum trees per acre on a condition 
					//for live trees >= 5 inches in diameter 
					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,tpacurr) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.tottpa as tpa  " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(tpacurr) as tottpa " + 
						"FROM " + this.m_strTreeTable + " " + 
						"WHERE  dia >= 5 AND statuscd=1 " + 
						"GROUP BY biosum_cond_id) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 2);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.TPACURR = u.TPACURR;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 3);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					//swd_tpa
					//sum trees per acre on a condition 
					//for softwood live trees >= 5 inches in diameter 
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,swd_tpacurr) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.totswdtpa as swd_tpa  " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(tpacurr) as totswdtpa " + 
						"FROM " + this.m_strTreeTable + " " + 
						"WHERE spcd < 300 AND dia >= 5 AND statuscd=1 " + 
						"GROUP BY biosum_cond_id ) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 4);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.SWD_TPACURR = u.SWD_TPACURR;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 5);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					//hwd tpa
					//sum trees per acre on a condition 
					//for hardwood live trees >= 5 inches in diameter 
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,hwd_tpacurr) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.tothwdtpa as hwd_tpacurr  " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(tpacurr) as tothwdtpa " + 
						"FROM " + this.m_strTreeTable + " " + 
						"WHERE spcd > 299 AND dia >= 5 AND statuscd=1 " + 
						"GROUP BY biosum_cond_id ) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 6);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.HWD_TPACURR = u.HWD_TPACURR;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 7);
                

                //vol_ac_grs_ft3
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					//total
					//for all live trees >= 5 inches in diameter 
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,vol_ac_grs_ft3) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.tot_volgrsft3 as vol_ac_grs_ft3  " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(volcfgrs * tpacurr) as tot_volgrsft3 " + 
						"FROM " + this.m_strTreeTable + " WHERE volcfgrs IS NOT NULL AND tpacurr IS NOT NULL AND statuscd=1 AND dia >=5 " + 
						"GROUP BY biosum_cond_id ) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 8);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.VOL_AC_GRS_FT3 = u.VOL_AC_GRS_FT3;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 9);

                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					//hwd
					//for all live hardwood trees >= 5 inches in diameter 
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,hwd_vol_ac_grs_ft3) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.tot_volgrsft3 as hwd_vol_ac_grs_ft3  " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(volcfgrs * tpacurr) as tot_volgrsft3 " + 
					 	 "FROM " + this.m_strTreeTable + " " + 
                         "WHERE spcd > 299 AND volcfgrs IS NOT NULL AND tpacurr IS NOT NULL AND statuscd=1 AND dia >=5 " + 
						"GROUP BY biosum_cond_id ) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 10);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.HWD_VOL_AC_GRS_FT3 = u.HWD_VOL_AC_GRS_FT3;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 11);



                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					//SWD
					//for all live softwood trees >= 5 inches in diameter 
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,swd_vol_ac_grs_ft3) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.tot_volgrsft3 as swd_vol_ac_grs_ft3  " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(volcfgrs * tpacurr) as tot_volgrsft3 " + 
						"FROM " + this.m_strTreeTable + " " + 
						"WHERE spcd < 300 AND volcfgrs IS NOT NULL AND tpacurr IS NOT NULL AND statuscd=1 AND dia >=5 " + 
						"GROUP BY biosum_cond_id ) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 12);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.SWD_VOL_AC_GRS_FT3 = u.SWD_VOL_AC_GRS_FT3;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value",13);

                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					//ba_ft2_ac basal area column
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,ba_ft2_ac) " + 
						"SELECT a.biosum_cond_id, b.tottemp AS ba_ft2_ac " + 
						"FROM " + this.m_strCondTable + " a, " +  
						"(SELECT biosum_cond_id, SUM((.005454154 * dia^2)  * tpacurr)  AS tottemp " + 
						 "FROM " + this.m_strTreeTable + " " +
                         "WHERE biosum_status_cd=9  AND statuscd=1 AND dia >= 1 " + 
						 "GROUP BY biosum_cond_id)  b " + 
						"WHERE a.biosum_cond_id = b.biosum_cond_id ";


					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 14);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.BA_FT2_AC = u.BA_FT2_AC;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 15);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					//swd_ba_ft2_ac softwood basal area 
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,swd_ba_ft2_ac) " + 
						"SELECT a.biosum_cond_id, b.tottemp AS swd_ba_ft2_ac " + 
						"FROM " + this.m_strCondTable + " a, " +  
						"(SELECT biosum_cond_id, SUM((.005454154 * dia^2)  * tpacurr)  AS tottemp " + 
						"FROM " + this.m_strTreeTable + " " + 
						"WHERE biosum_status_cd=9 AND " +
                              "spcd < 300 AND statuscd=1 AND dia >= 1 " + 
						"GROUP BY biosum_cond_id)  b " + 
						"WHERE a.biosum_cond_id = b.biosum_cond_id ";


					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 16);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.SWD_BA_FT2_AC = u.SWD_BA_FT2_AC";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 17);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					//hardwood ba_ft2_ac
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,hwd_ba_ft2_ac) " + 
						"SELECT a.biosum_cond_id, b.tottemp AS hwd_ba_ft2_ac " + 
						"FROM " + this.m_strCondTable + " a, " +  
						"(SELECT biosum_cond_id, SUM((.005454154 * dia^2)  * tpacurr)  AS tottemp " + 
						"FROM " + this.m_strTreeTable + " " + 
						"WHERE biosum_status_cd=9 AND " +
                        "spcd > 299 AND statuscd=1 AND dia >= 1 " + 
						"GROUP BY biosum_cond_id)  b " + 
						"WHERE a.biosum_cond_id = b.biosum_cond_id ";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 18);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.HWD_BA_FT2_AC = u.HWD_BA_FT2_AC";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 19);

				//volcsgrs  
				//gross sawlog
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,volcsgrs) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.ttl as  volcsgrs " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(volcsgrs) as ttl " + 
						"FROM " + this.m_strTreeTable + " " + 
						"GROUP BY biosum_cond_id) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 20);

                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.volcsgrs = u.volcsgrs;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);

				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 21);


				//swd_volcsgrs      
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,swd_volcsgrs) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.ttl as swd_volcsgrs  " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(volcsgrs) as ttl " + 
						"FROM " + this.m_strTreeTable + " " + 
						"WHERE SPCD < 300 " + 
						"GROUP BY biosum_cond_id ) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 22);

                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.SWD_volcsgrs = u.SWD_volcsgrs";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 23);

				//hwd_volcsgrs      
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,hwd_volcsgrs) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.ttl as hwd_volcsgrs  " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(volcsgrs) as ttl " + 
						"FROM " + this.m_strTreeTable + " " + 
						"WHERE SPCD > 299 " + 
						"GROUP BY biosum_cond_id ) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 24);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.HWD_volcsgrs = u.HWD_volcsgrs";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 25);








				//qmd_tot_cm 
				// quadratic mean diameter for all the live trees on the condition
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,qmd_tot_cm) " + 
						"SELECT c.biosum_cond_id, SQR(c.ba_ft2_ac/(.005454154 * c.tpacurr)) as qmd_tot_cm " + 
						"FROM " + this.m_strCondTable + " c " + 
						"WHERE c.biosum_status_cd=9 AND " + 
						      "c.ba_ft2_ac IS NOT NULL AND " + 
							  "c.ba_ft2_ac <> 0 AND " + 
						      "c.tpacurr IS NOT NULL AND " + 
						      "c.tpacurr <> 0;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 26);

                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.qmd_tot_cm = u.qmd_tot_cm;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);

				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 27);


				//swd_qmd_tot_cm      
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,swd_qmd_tot_cm) " + 
						"SELECT c.biosum_cond_id, SQR(c.swd_ba_ft2_ac/(.005454154 * c.swd_tpacurr)) as swd_qmd_tot_cm " + 
						"FROM " + this.m_strCondTable + " c " + 
						"WHERE c.biosum_status_cd=9 AND " + 
						"c.swd_ba_ft2_ac IS NOT NULL AND " + 
						"c.swd_ba_ft2_ac <> 0 AND " + 
						"c.swd_tpacurr IS NOT NULL AND " + 
						"c.swd_tpacurr <> 0;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 28);

                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.SWD_qmd_tot_cm = u.SWD_qmd_tot_cm";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 29);

				//hwd_qmd_tot_cm    
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,hwd_qmd_tot_cm) " + 
						"SELECT c.biosum_cond_id, SQR(c.hwd_ba_ft2_ac/(.005454154 * c.hwd_tpacurr)) as hwd_qmd_tot_cm " + 
						"FROM " + this.m_strCondTable + " c " + 
						"WHERE c.biosum_status_cd=9 AND " + 
						"c.hwd_ba_ft2_ac IS NOT NULL AND " + 
						"c.hwd_ba_ft2_ac <> 0 AND " + 
						"c.hwd_tpacurr IS NOT NULL AND " + 
						"c.hwd_tpacurr <> 0;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 30);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.HWD_qmd_tot_cm = u.HWD_qmd_tot_cm";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 31);



				//VOL_AC_GRS_STEM_TTL_FT
                //gross wood volume of the total stem from ground to tip
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);


					
					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,vol_ac_grs_stem_ttl_ft3) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.ttl as  vol_ac_grs_stem_ttl_ft3 " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(IIF(dia >= 5, (drybiot / (drybiom/volcfgrs)) * tpacurr,IIF(spcd < 300, (drybiot /25.82) * tpacurr,(drybiot /31.79) * tpacurr))) AS ttl " + 
						"FROM " + this.m_strTreeTable + " " + 
						"WHERE statuscd=1 AND dia >= 1 " + 
						"GROUP BY biosum_cond_id) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 32);

                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.vol_ac_grs_stem_ttl_ft3 = u.vol_ac_grs_stem_ttl_ft3;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);

				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 33);


				//hwd_vol_ac_grs_stem_ttl_ft
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,hwd_vol_ac_grs_stem_ttl_ft3) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.ttl as  hwd_vol_ac_grs_stem_ttl_ft3 " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id,  SUM(IIF(dia >= 5, (drybiot / (drybiom/volcfgrs)) * tpacurr,(drybiot /31.79) * tpacurr)) as ttl " + 
						"FROM " + this.m_strTreeTable + " " + 
						"WHERE statuscd=1 AND dia >= 1 AND spcd > 299 " + 
						"GROUP BY biosum_cond_id) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";


					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 34);

                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.hwd_vol_ac_grs_stem_ttl_ft3 = u.hwd_vol_ac_grs_stem_ttl_ft3";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 35);

				//swd_vol_ac_grs_stem_ttl_ft     
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,swd_vol_ac_grs_stem_ttl_ft3) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.ttl as  swd_vol_ac_grs_stem_ttl_ft3 " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id,  SUM( IIF(dia >= 5, (drybiot / (drybiom/volcfgrs)) * tpacurr,(drybiot /25.82) * tpacurr)) as ttl  " + 
						"FROM " + this.m_strTreeTable + " " + 
						"WHERE statuscd=1 AND dia >= 1 AND spcd < 300 " + 
						"GROUP BY biosum_cond_id) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";


					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 36);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.swd_vol_ac_grs_stem_ttl_ft3 = u.swd_vol_ac_grs_stem_ttl_ft3";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 37);
                //
                //VOLTSGRS column update
                //
                SetLabelValue(m_frmTherm.lblMsg, "Text", "Start Oracle Services...Stand By");
                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
                FIADB.Oracle.Services m_oOracleServices = new FIADB.Oracle.Services();
                m_oOracleServices.Start();
                SetThermValue(m_frmTherm.progressBar1, "Value", 38);

                if (m_oOracleServices.m_oTree == null) MessageBox.Show("m_oTree==null");
                m_oOracleServices.m_oTree.GetVolumesMode = FIADB.Oracle.Services.Tree.GetVolumesModeValues.SQLUpdate;
                //if (m_strGridTableSource.Trim() != Tables.FVS.DefaultOracleInputVolumesTable)
                //{
                    //step 5 - delete and create work tables
                    if (p_ado.TableExist(this.m_connTempMDBFile, Tables.FVS.DefaultOracleInputVolumesTable))
                        p_ado.SqlNonQuery(this.m_connTempMDBFile, "DROP TABLE " + Tables.FVS.DefaultOracleInputVolumesTable);
                    frmMain.g_oTables.m_oFvs.CreateOracleInputBiosumVolumesTable(p_ado, this.m_connTempMDBFile, Tables.FVS.DefaultOracleInputVolumesTable);

                    if (p_ado.TableExist(this.m_connTempMDBFile, Tables.FVS.DefaultOracleInputFCSVolumesTable))
                        p_ado.SqlNonQuery(this.m_connTempMDBFile, "DROP TABLE " + Tables.FVS.DefaultOracleInputFCSVolumesTable);
                    frmMain.g_oTables.m_oFvs.CreateOracleInputFCSBiosumVolumesTable(p_ado, this.m_connTempMDBFile, Tables.FVS.DefaultOracleInputFCSVolumesTable);

                   

                    //step 6 - insert records
                    strColumns = "STATECD,COUNTYCD,PLOT,INVYR,TREE,SPCD,DIA,HT," +
                                "ACTUALHT,CR,STATUSCD,TREECLCD,ROUGHCULL,CULL,TRE_CN,CND_CN,PLT_CN,VOL_LOC_GRP";


                    strValues = "STATECD," +
                                "COUNTYCD," +
                                "CINT(MID(BIOSUM_COND_ID,16,5)) AS PLOT," +
                                "INVYR,TREE,SPCD,IIF(DIA IS NOT NULL,ROUND(DIA,2),DIA),HT,ACTUALHT,CR,STATUSCD,TREECLCD,ROUGHCULL,CULL," +
                                "CN AS TRE_CN," +
                                "BIOSUM_COND_ID AS CND_CN," +
                                "MID(BIOSUM_COND_ID,1,LEN(BIOSUM_COND_ID)-1) AS PLT_CN,'' AS VOL_LOC_GRP";

                    p_ado.m_strSQL = "INSERT INTO " + Tables.FVS.DefaultOracleInputFCSVolumesTable + " " +
                                     "(" + strColumns + ") SELECT " + strValues + " FROM " + m_strTreeTable;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n\r\n");
                    p_ado.SqlNonQuery(this.m_connTempMDBFile, p_ado.m_strSQL);
              
                    p_ado.m_strSQL = "UPDATE " + Tables.FVS.DefaultOracleInputFCSVolumesTable + " f INNER JOIN " + m_strCondTable + " c ON f.CND_CN = c.BIOSUM_COND_ID SET f.vol_loc_grp=IIF(INSTR(1,c.vol_loc_grp,'22') > 0,'S26LEOR',c.vol_loc_grp)";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n\r\n");
                    p_ado.SqlNonQuery(this.m_connTempMDBFile, p_ado.m_strSQL);


                p_ado.m_strSQL = "INSERT INTO fcs_biosum_volume (" + strColumns + ") SELECT " + strColumns + " FROM " + Tables.FVS.DefaultOracleInputFCSVolumesTable;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,p_ado.m_strSQL + "\r\n\r\n");
                p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
                SetThermValue(m_frmTherm.progressBar1, "Value", 39);
               
               
                SetLabelValue(m_frmTherm.lblMsg, "Text", "Wait For Oracle Volume Compilation To Complete...Stand By");
                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
               
                m_oOracleServices.m_oTree.GetBiosumVolumes();

                SetThermValue(m_frmTherm.progressBar1, "Value", 40);

                if (m_oOracleServices.m_intError == 0)
                {
                    SetLabelValue(m_frmTherm.lblMsg, "Text", "Update tree VOLTSGRS column with Oracle Calculated Values...Stand By");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
                    string strConn = m_connTempMDBFile.ConnectionString;
                    p_ado.CloseConnection(m_connTempMDBFile);
                    p_ado.OpenConnection(strConn,ref m_connTempMDBFile);
                    p_ado.m_strSQL = "UPDATE " + m_strTreeTable + " t " +
                                     "INNER JOIN fcs_biosum_volume f " +
                                     "ON t.cn = f.tre_cn " +
                                     "SET t.VOLTSGRS = f.VOLTSGRS_CALC";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n\r\n");
                    p_ado.SqlNonQuery(m_connTempMDBFile, p_ado.m_strSQL);
                   
                }
                SetThermValue(m_frmTherm.progressBar1, "Value", 41);
  

			}
       
			

           
			//----------------------PLOT COLUMN UPDATES-----------------------//
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//----------------------PLOT COLUMN UPDATES-----------------------//\r\n");
            SetThermValue(m_frmTherm.progressBar1, "Maximum", 3);
            SetThermValue(m_frmTherm.progressBar1, "Minimum", 0);
            SetThermValue(m_frmTherm.progressBar1, "Value", 0);
            SetLabelValue(m_frmTherm.lblMsg,"Text","Updating Plot Table Columns...Stand By");
            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

            if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				/********************************************
				 **update gis_status_id for idb records
				 ********************************************/
				if (Checked(rdoIDB)==true)
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strPlotTable + " p " + 
						             " SET p.gis_status_id = 1 " + 
						             " WHERE biosum_status_cd = 9;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					/**************************************************
					**update plot_accessible_yn
					***************************************************/
					p_ado.m_strSQL = "UPDATE " + this.m_strPlotTable + " p " + 
						             " SET p.plot_accessible_yn = 'Y', " +
						                  "p.gis_protected_area_yn = 'N'," + 
						                  "p.gis_roadless_yn = 'N'," + 
						                  "p.all_cond_not_accessible_yn='N' " + 
						             " WHERE biosum_status_cd = 9;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
				}
				/********************************************
				 **update the plot half state field
				 ********************************************/
				p_ado.m_strSQL = "UPDATE " + this.m_strPlotTable + " p " + 
					" INNER JOIN " + this.m_strCondTable + " c " + 
					"ON p.biosum_plot_id = c.biosum_plot_id  " + 
					" SET p.half_state = MID(c.vol_loc_grp,5,LEN(TRIM(c.vol_loc_grp))) " + 
					" WHERE c.condid=1;";

				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
				p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
				//MessageBox.Show(strTime);

				
				 
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 1);

            if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				/***************************************************
				 **update the number of conditions on each plot
				 ***************************************************/
				//use the biosum_plot_input as our work table so delete all records
				p_ado.m_strSQL = "DELETE FROM plot_column_updates_work_table ";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
				p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

				//insert the condition counts into the work table
				p_ado.m_strSQL = "INSERT INTO plot_column_updates_work_table (biosum_plot_id, cond_ttl) " + 
					" SELECT biosum_plot_id , COUNT(biosum_plot_id) " + 
					" FROM " + this.m_strCondTable + 
					" GROUP BY biosum_plot_id;";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
				p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 2);

            if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{

				p_ado.m_strSQL = "UPDATE " + this.m_strPlotTable  + " p " + 
					"INNER JOIN plot_column_updates_work_table i " + 
					"ON  p.biosum_plot_id = i.biosum_plot_id " + 
					"SET p.num_cond = i.cond_ttl, p.one_cond_yn = IIF(i.cond_ttl > 1,'N','Y');";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
				p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 3);





            this.m_intError=p_ado.m_intError;







			//MessageBox.Show(strTime);
			
		}
	
		private void ThermCancel(object sender, System.EventArgs e)
		{
			string strMsg = "Do you wish to cancel appending plot data (Y/N)?";
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
		/// <summary>
		/// create a delimited string list from a text file
		/// that has a single column of data with multiple rows
		/// </summary>
		/// <param name="p_strTxtFile">text file containing the column of data</param>
		/// <param name="p_strTxtFileDelimiter">specified character between list items</param>
		/// <param name="p_strListDelimiter">specified character between list items</param>
		/// <param name="p_bNumericDataType">specifies if the column data to retrieve in the text file is numeric</param>
		/// <returns></returns>
		private string CreateDelimitedStringList(string p_strTxtFile,string p_strTxtFileDelimiter, string p_strListDelimiter,bool p_bNumericDataType)
		{
			//The DataSet to Return
			//DataSet result = new DataSet();
			this.m_intError=0;
			string strList="";
			string str="";
			try
			{
				//Open the file in a stream reader.
				System.IO.StreamReader s = new System.IO.StreamReader(p_strTxtFile);
				//Read the rest of the data in the file.        
				string AllData = s.ReadToEnd();
    
				//Split off each row at the Carriage Return/Line Feed
				//Default line ending in most <A class=iAs style="FONT-WEIGHT: normal; FONT-SIZE: 100%; PADDING-BOTTOM: 1px; COLOR: darkgreen; BORDER-BOTTOM: darkgreen 0.07em solid; BACKGROUND-COLOR: transparent; TEXT-DECORATION: underline" href="#" target=_blank itxtdid="2592535">windows</A> exports.  
				string[] rows = AllData.Split("\r\n".ToCharArray());
 
				//Now add each row to the DataSet        
				foreach(string r in rows)
				{
					//Split the row at the delimiter.
					string[] items = r.Split(p_strTxtFileDelimiter.ToCharArray());
					str = items[0].Trim();  //plot_cn in first column
					str = str.Replace("\"",""); //remove any quotations
					if (str.Trim().Length > 0)
					{
						if (strList.Trim().Length == 0)
						{
							if (p_bNumericDataType == true)
							{
								strList = str.Trim();
							}
							else
							{
								strList = "'" + str.Trim() + "'";
							}
						}
						else
						{
							if (p_bNumericDataType == true)
							{
								strList = strList + p_strListDelimiter.Trim() + str.Trim();
							}
							else
							{
								strList = strList + p_strListDelimiter.Trim() + "'" + str.Trim() + "'";
							}

						}
					}
				}
			}
			catch (Exception caught)
			{
				this.m_intError=-1;
				MessageBox.Show("!!Error: CreateDelimitedStringList() Routine Error Msg:" + caught.Message);
			}
			return strList;
		}

		private void btnFilterByFileBrowse_Click(object sender, System.EventArgs e)
		{
			
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			if (this.rdoFIADB.Checked==true)
			{
				OpenFileDialog1.Title = "Text File With PLOT_CN data";
			}
			else
			{
				OpenFileDialog1.Title = "Text File With IDB_PLOT_ID data";
			}
			OpenFileDialog1.Filter = "Text File (*.TXT;*.DAT) |*.txt;*.dat";
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK) 
			{
				if (OpenFileDialog1.FileName.Trim().Length > 0) 
				{
					this.txtFilterByFile.Text = OpenFileDialog1.FileName.Trim();
				}
				else 
				{
				}
				OpenFileDialog1 = null;

			}
		}

		private void rdoIDB_Click(object sender, System.EventArgs e)
		{
			if (rdoIDB.Checked==true)
			{
				this.grpInputDataSourceType.Enabled=false;
			}
		}

		private void rdoFIADB_Click(object sender, System.EventArgs e)
		{
			if (rdoFIADB.Checked==true)
			{
				this.grpInputDataSourceType.Enabled=true;
			}

		}

		private void btnFilterNext_Click(object sender, System.EventArgs e)
		{
			if (this.rdoFIADB.Checked==true)
			{
				//
				//see if text file input
				//
				
				if (this.rdoAccess.Checked)
				{
					if (this.LoadMDBFiadbPopEvalTable() && m_intError==0)
					{	
						this.m_strLoadedPopEvalInputTable=this.cmbFiadbPopEvalTable.Text;
						this.FIADBLoadInv();
						
					}
					else if (m_intError==0)
					{
					
					}
				}
				else
				{
					//see if current pop eval id text file is loaded
					if (this.m_strPopEvalTxtInputFile.Trim().ToUpper() !=
						this.m_strLoadedPopEvalTxtInputFile.Trim().ToUpper())
					{
						//input the pop eval table to 
						//identify inventories 
						this.m_strTableType="POPULATION EVALUATION";
						this.m_strCurrentTxtInputFile = this.txtPopEval.Text;
						this.StartTherm("1","Add Population Evaluation Data");
						this.thdProcessRecords = new Thread(new ThreadStart(this.txtFileInputPopFiles));
						this.thdProcessRecords.IsBackground = true;
						this.thdProcessRecords.Start();
						while (thdProcessRecords.IsAlive)
						{
							thdProcessRecords.Join(1000);
							System.Windows.Forms.Application.DoEvents();

						}
						thdProcessRecords=null;
						this.m_frmTherm.Close();
						this.m_frmTherm = null;
						if (m_intError==0)
						{
							this.m_strLoadedPopEvalTxtInputFile=this.m_strCurrentTxtInputFile;
							FIADBLoadInv();
						}
					}
						
					else
					{
					
					
					}
				}
				if (this.m_intError==0)
				{
					if (this.rdoFilterByMenu.Checked==true)
					{
						this.btnFIADBInvAppend.Enabled=false;
						this.btnFIADBInvNext.Enabled=true;
					}
					else
					{
						this.btnFIADBInvAppend.Enabled=true;
						this.btnFIADBInvNext.Enabled=false;
					}
					this.grpboxFIADBInv.Visible=true;
					this.grpboxFilter.Visible=false;
				}
				
				//}
				//
				//see if MSAccess input
				//
			}
			else if (this.rdoAccess.Checked)
			{
			}


			
			if (this.rdoIDB.Checked==true)
			{
				if (this.rdoFilterByMenu.Checked==true)
				{
					this.btnIDBInvAppend.Enabled=false;
					this.btnIDBInvNext.Enabled=true;
				}
				else
				{
					this.btnIDBInvAppend.Enabled=true;
					this.btnIDBInvNext.Enabled=false;
				}
				this.idbLoadInv();
				if (this.m_intError==0)
				{
					this.grpboxIDBInv.Visible=true;
					this.grpboxFilter.Visible=false;
				}
			}
		}
		private void txtFileInputStateCounty()
		{
			//see if we have already loaded the list box with the current inventory
			if (this.m_bLoadStateCountyList==false && this.lstFilterByState.Items.Count > 0) return;

			string strState="";
			string strCounty="";
			string strPlot="";
			string strCn="";
			int intAddedPlotRows=0;
			bool bAddRow=true;
			int x=0;
			
			int intPlotStatusCd=0;
			string strKey;
			System.Data.DataRow  p_rowFound;
			System.Data.DataRow  p_rowAdd;
			System.Data.DataRow  p_rowAdd2;
               
			
			//instatiate the oledb data access class
			this.m_ado = new ado_data_access();

			//load up the current population plot stratum assignment plots for the selected evaluation
			System.Data.DataTable dtPlotCN = new DataTable("PlotCN");
			dtPlotCN.Columns.Add("plt_cn",typeof(string));
			// two columns in the Primary Key.
			DataColumn[] colPk = new DataColumn[1];
			colPk[0] =dtPlotCN.Columns["plt_cn"];
            dtPlotCN.PrimaryKey=colPk;
			this.m_ado.SqlQueryReader(this.m_strTempMDBFileConn,"SELECT DISTINCT plt_cn FROM " + this.m_strPpsaTable + " " + 
															    "WHERE rscd=" + this.m_strCurrFIADBRsCd + " AND " + 
																	   "evalid=" + this.m_strCurrFIADBEvalId + " AND " + 
																		"biosum_status_cd=9");
			if (this.m_ado.m_OleDbDataReader.HasRows)
			{
				while (this.m_ado.m_OleDbDataReader.Read())
				{
					if (this.m_ado.m_OleDbDataReader["plt_cn"] != System.DBNull.Value)
					{
						strCn = Convert.ToString(this.m_ado.m_OleDbDataReader["plt_cn"]);
						if (strCn.Trim().Length > 0)
						{
							p_rowAdd = dtPlotCN.NewRow();
							p_rowAdd["plt_cn"] = strCn.Trim();
							dtPlotCN.Rows.Add(p_rowAdd);	
						}
					}
				}
			}
			this.m_ado.m_OleDbDataReader.Close();
				    

			this.m_dtStateCounty.Clear();        
			this.lstFilterByState.Clear();
			this.lstFilterByState.Columns.Add(" ", 100, HorizontalAlignment.Center); 
			this.lstFilterByState.Columns.Add("State", 100, HorizontalAlignment.Left);
			this.lstFilterByState.Columns.Add("County", 100, HorizontalAlignment.Left);

			this.m_strStateCountyPlotSQL="";
			this.m_strStateCountySQL="";
			this.m_intError=0;

			
			this.m_frmTherm = new frmTherm();
			this.m_frmTherm.btnCancel.Click += new System.EventHandler(this.ThermCancel);
			this.m_frmTherm.Text = "Load State, County, Plot Menus";
			this.m_frmTherm.Visible=false;
			this.m_frmTherm.btnCancel.Visible=true;
			this.m_frmTherm.lblMsg.Visible=true;
			this.m_frmTherm.progressBar1.Minimum=0;
			this.m_frmTherm.progressBar1.Visible=true;
			this.m_frmTherm.progressBar1.Maximum = 10;
			this.m_frmTherm.AbortProcess = false;
			this.m_frmTherm.Refresh();
			this.m_frmTherm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			((frmDialog)this.ParentForm).Enabled=false;
					

			FIA_Biosum_Manager.env p_env = new env();
			this.m_ado.m_DataSet = new DataSet("FIADB");
			this.m_ado.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();

			//----------------PLOT DATA---------------//
			this.m_ado.ConvertDelimitedTextToDataTable(this.m_ado.m_DataSet,this.txtPlot.Text.Trim(),"fiadb_plot",",");
			if (this.m_ado.m_intError==0)
			{
				try
				{
					if (this.m_ado.m_DataSet.Tables["fiadb_plot"].Rows.Count > 0)
					{
								
						this.m_frmTherm.progressBar1.Minimum=0;
						this.m_frmTherm.progressBar1.Maximum = this.m_ado.m_DataSet.Tables["fiadb_plot"].Rows.Count-1;
						this.m_frmTherm.lblMsg.Text="Plot Table: State, County, Plot Records";
						this.m_frmTherm.Visible=true;
						this.m_frmTherm.Refresh();
						//load up each row in the FIADB plot input table
						for (x = 0; x<=this.m_ado.m_DataSet.Tables["fiadb_plot"].Rows.Count-1;x++)
						{
							strState="";
							strCounty="";
							strPlot="";
							strCn="";
							bAddRow=true;
							this.m_frmTherm.progressBar1.Value=x;
							//make sure the row is not null values
							if (this.m_ado.m_DataSet.Tables["fiadb_plot"].Rows[x][0] != System.DBNull.Value &&
								this.m_ado.m_DataSet.Tables["fiadb_plot"].Rows[x][0].ToString().Trim().Length > 0)
							{
								strCn   = this.m_ado.m_DataSet.Tables["fiadb_plot"].Rows[x]["cn"].ToString().Trim();
								//lookup the cn in the ppsa table
								strKey = strCn;
								System.Object[] p_searchCN = new Object[1];
								p_searchCN[0] = strCn.Trim();
								p_rowFound = dtPlotCN.Rows.Find(p_searchCN);
								if (p_rowFound != null)
								{
									strState= this.m_ado.m_DataSet.Tables["fiadb_plot"].Rows[x]["statecd"].ToString();
									strCounty = this.m_ado.m_DataSet.Tables["fiadb_plot"].Rows[x]["countycd"].ToString();
									strPlot = this.m_ado.m_DataSet.Tables["fiadb_plot"].Rows[x]["plot"].ToString().Trim();
								

									if (this.rdoFilterNone.Checked==true || 
										this.rdoFilterByMenu.Checked==true)
									{
										intPlotStatusCd=1;
										if (this.m_ado.m_DataSet.Tables["fiadb_plot"].Rows[x]["plot_status_cd"] != System.DBNull.Value)
										{
											intPlotStatusCd=Convert.ToInt32(this.m_ado.m_DataSet.Tables["fiadb_plot"].Rows[x]["plot_status_cd"]);
										}

										if (this.chkForested.Checked==true &&
											this.chkNonForested.Checked==false)								
										{
											if (this.m_ado.m_DataSet.Tables["fiadb_plot"].Rows[x]["plot_status_cd"] != System.DBNull.Value)
											{
												if (intPlotStatusCd != 1) 
												{
													bAddRow=false;
												}
											}
										}
										else if (this.chkForested.Checked==false &&
											this.chkNonForested.Checked==true)
										{
											if (intPlotStatusCd == 1)
											{
												bAddRow=false;
											}
										}
									}

									if (bAddRow==true)
									{
										strKey = strState + strCounty;
										//see if the state and county have been loaded already
										System.Object[] p_search = new Object[2];
										p_search[0] = strState.Trim();
										p_search[1] = strCounty.Trim();

										p_rowFound = this.m_dtStateCounty.Rows.Find(p_search);
										if (p_rowFound == null)
										{
											//save the state,county combination
											p_rowAdd = this.m_dtStateCounty.NewRow();
											p_rowAdd["statecd"] = strState.Trim();
											p_rowAdd["countycd"] = strCounty.Trim();
											this.m_dtStateCounty.Rows.Add(p_rowAdd);
											this.lstFilterByState.BeginUpdate();
											System.Windows.Forms.ListViewItem listItem = new ListViewItem();
											listItem.Checked=false;
											listItem.SubItems.Add(strState);
											listItem.SubItems.Add(strCounty);
											this.lstFilterByState.Items.Add(listItem);
											this.lstFilterByState.EndUpdate();
										}
										p_rowAdd2 = this.m_dtPlot.NewRow();
										p_rowAdd2["statecd"] = strState.Trim();
										p_rowAdd2["countycd"] = strCounty.Trim();
										p_rowAdd2["plot"] = strPlot;
										this.m_dtPlot.Rows.Add(p_rowAdd2);
										intAddedPlotRows++;
									}
								}
								
								System.Windows.Forms.Application.DoEvents();
								if (this.m_frmTherm.AbortProcess == true) break;
								
							}
							
						}
						if (intAddedPlotRows == 0 && this.m_frmTherm.AbortProcess==false)
						{
							this.m_intError=-1;
							MessageBox.Show("!!No Plots Loaded To Get State, County, Plot Information!!","Load State, County, Plot Menus");
						}
					}
					else
					{
						this.m_intError=-1;
						MessageBox.Show("!!No Plot Records In The FIADB Plot Input Table!!");
					}
				}
				catch (Exception caught)
				{
					this.m_intError=-1;
					MessageBox.Show(caught.Message);
				}
			}
			
			if (this.m_intError==0) this.m_bLoadStateCountyList=false;
			
			this.m_frmTherm.Close();
			this.m_frmTherm = null;
			dtPlotCN.Clear();
			dtPlotCN.Dispose();
			dtPlotCN=null;
			((frmDialog)this.ParentForm).Enabled=true;
		}
		private void txtFileInputPlot()
		{
			//see if already loaded
			if (this.m_bLoadStateCountyPlotList==false && this.lstFilterByPlot.Items.Count > 0) return;

			this.m_intError=0;
			
			string strState="";
			string strCounty="";
			string strPlot="";
			this.m_intError=0;
			System.Data.DataRow[] p_rows;
			this.lstFilterByPlot.Clear();
			this.lstFilterByPlot.Columns.Add(" ", 50, HorizontalAlignment.Center); 
			this.lstFilterByPlot.Columns.Add("State", 75, HorizontalAlignment.Left);
			this.lstFilterByPlot.Columns.Add("County", 75, HorizontalAlignment.Left);
			this.lstFilterByPlot.Columns.Add("Plot", 100, HorizontalAlignment.Left);

			for (int x=0;x<=this.lstFilterByState.Items.Count-1;x++)
			{
				if (this.lstFilterByState.Items[x].Checked==true)
				{
					strState = this.lstFilterByState.Items[x].SubItems[1].Text.Trim();
					strCounty = this.lstFilterByState.Items[x].SubItems[2].Text.Trim();
					p_rows = this.m_dtPlot.Select("trim(statecd) = '" + strState.Trim()  + "' and trim(countycd) = '" + strCounty.Trim() + "'");
					if (p_rows != null)
					{
						for (int y=0; y <= p_rows.Length-1;y++)
						{
							strPlot = p_rows[y]["plot"].ToString().Trim();
							this.lstFilterByPlot.BeginUpdate();
							System.Windows.Forms.ListViewItem listItem = new ListViewItem();
							listItem.Checked=false;
							listItem.SubItems.Add(strState);
							listItem.SubItems.Add(strCounty);
							listItem.SubItems.Add(strPlot);
							this.lstFilterByPlot.Items.Add(listItem);
							this.lstFilterByPlot.EndUpdate();
						}
						
					}
                    
				}
			}
			if (strPlot.Trim().Length ==0)
			{
				this.m_intError=-1;
				MessageBox.Show("Select A State And County","Add Plot Data", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			}
			if (this.m_intError==0) this.m_bLoadStateCountyPlotList=false;



		}

		private void mdbInputStateCounty()
		{
	
			string strState="";
			string strCounty="";
			int intAddedPlotRows=0;
    

			this.m_dtStateCounty.Clear();        
			this.lstFilterByState.Clear();
			this.lstFilterByState.Columns.Add(" ", 100, HorizontalAlignment.Center); 
			this.lstFilterByState.Columns.Add("State", 100, HorizontalAlignment.Left);
			this.lstFilterByState.Columns.Add("County", 100, HorizontalAlignment.Left);

			this.m_strStateCountyPlotSQL="";
			this.m_strStateCountySQL="";
			this.m_intError=0;

			FIA_Biosum_Manager.ado_data_access p_ado = new ado_data_access();

			string strMDBFile;
			string strTable;
			string strConn;
			if (this.rdoFIADB.Checked==false)
			{
				strMDBFile = this.txtMDBPlot.Text.Trim();
				strTable = this.txtMDBPlotTable.Text.Trim();
				strConn = p_ado.getMDBConnString(strMDBFile,"","");
			}
			else
			{
				strMDBFile = this.txtMDBFiadbInputFile.Text.Trim();
				strTable = this.cmbFiadbPpsaTable.Text.Trim();
				strConn = p_ado.getMDBConnString(strMDBFile,"","");
			}


			
			


			if (this.chkNonForested.Checked == true && this.chkForested.Checked==true)
			{
				if (rdoFIADB.Checked)
				{
					p_ado.m_strSQL = "SELECT statecd,countycd " + 
						"FROM " + strTable +  " " + 
						"WHERE RSCD = " + this.m_strCurrFIADBRsCd + " AND " + 
						"EVALID = " + this.m_strCurrFIADBEvalId + " " +  
						"GROUP BY statecd,countycd;";
					
				}
				else
				{
					p_ado.m_strSQL = "SELECT statecd,countycd " + 
						"FROM " + strTable + " " + 
						"WHERE mid(biosum_plot_id,2,4) = '" + this.m_strIDBInv.Trim() + "' " + 
						"GROUP BY statecd,countycd;";
				}
			}
			else if (this.chkForested.Checked==true)
			{
				if (rdoFIADB.Checked)
				{
					p_ado.m_strSQL = "SELECT ppsa.statecd,ppsa.countycd " + 
						"FROM " + strTable +  " ppsa " + 
						"INNER JOIN " + this.cmbFiadbPlotTable.Text.Trim() + " p " + 
						"ON ppsa.plt_cn = p.cn " + 
						"WHERE ppsa.RSCD = " + this.m_strCurrFIADBRsCd + " AND " + 
						"ppsa.EVALID = " + this.m_strCurrFIADBEvalId + " AND " +  
						"p.plot_status_cd=1 " + 
						"GROUP BY ppsa.statecd,ppsa.countycd;";

				}
				else
				{
					p_ado.m_strSQL = "SELECT statecd,countycd " + 
						"FROM " + strTable + " " + 
						"WHERE plot_status_cd = 1 AND " + 
						"mid(biosum_plot_id,2,4) = '" + this.m_strIDBInv.Trim() + "' " + 
						"GROUP BY statecd,countycd;";

				}
			}
			else if (this.chkNonForested.Checked==true)
			{
				if (rdoFIADB.Checked)
				{
					p_ado.m_strSQL = "SELECT ppsa.statecd,ppsa.countycd " + 
						"FROM " + strTable +  " ppsa " + 
						"INNER JOIN " + this.cmbFiadbPlotTable.Text.Trim() + " p " + 
						"ON ppsa.plt_cn = p.cn " + 
						"WHERE ppsa.RSCD = " + this.m_strCurrFIADBRsCd + " AND " + 
						"ppsa.EVALID = " + this.m_strCurrFIADBEvalId + " AND " +  
						"p.plot_status_cd<>1 " + 
						"GROUP BY ppsa.statecd,ppsa.countycd;";
				}
				else
				{
					p_ado.m_strSQL = "SELECT statecd,countycd " + 
						"FROM " + strTable + " " + 
						"WHERE plot_status_cd <> 1 AND " + 
						"mid(biosum_plot_id,2,4) = '" + this.m_strIDBInv.Trim() + "' " + 
						"GROUP BY statecd,countycd;";
				}
			}
			else
			{
				this.m_intError=-1;

			}


			

           
			if (p_ado.m_intError==0 && this.m_intError==0)
			{
				p_ado.SqlQueryReader(strConn,p_ado.m_strSQL);
				if (p_ado.m_intError==0)
				{
					try
					{
						//load up each row in the FIADB plot input table

						while (p_ado.m_OleDbDataReader.Read())
						{
							strState="";
							strCounty="";
							
							//make sure the row is not null values
							if (p_ado.m_OleDbDataReader[0] != System.DBNull.Value &&
								p_ado.m_OleDbDataReader[0].ToString().Trim().Length > 0)
							{
								strState= p_ado.m_OleDbDataReader["statecd"].ToString();
								strCounty = p_ado.m_OleDbDataReader["countycd"].ToString();
								this.lstFilterByState.BeginUpdate();
								System.Windows.Forms.ListViewItem listItem = new ListViewItem();
								listItem.Checked=false;
								listItem.SubItems.Add(strState);
								listItem.SubItems.Add(strCounty);
								this.lstFilterByState.Items.Add(listItem);
								this.lstFilterByState.EndUpdate();
								intAddedPlotRows++;
							}

						}
						p_ado.m_OleDbDataReader.Close();
						if (intAddedPlotRows == 0 )
						{
							this.m_intError=-1;
							MessageBox.Show("!!No Plots Loaded To Get State, County, Plot Information!!","Load State, County, Plot Menus");
						}
						((frmDialog)this.ParentForm).Enabled=true;
					}
					catch (Exception caught)
					{
						this.m_intError=-1;
						MessageBox.Show(caught.Message);
					}
					p_ado.m_OleDbConnection.Close();
				}
				else
				{
					this.m_intError=p_ado.m_intError;
				}
			}
			
			
			p_ado=null;
		}

		private void mdbInputPlot()
		{
			this.m_intError=0;
			int intAddedPlotRows=0;
			string strState="";
			string strCounty="";
			string strPlot="";
			this.lstFilterByPlot.Clear();
			this.lstFilterByPlot.Columns.Add(" ", 50, HorizontalAlignment.Center); 
			this.lstFilterByPlot.Columns.Add("State", 75, HorizontalAlignment.Left);
			this.lstFilterByPlot.Columns.Add("County", 75, HorizontalAlignment.Left);
			this.lstFilterByPlot.Columns.Add("Plot", 100, HorizontalAlignment.Left);


			

			FIA_Biosum_Manager.ado_data_access p_ado = new ado_data_access();

			
			string strMDBFile;
			string strTable;
			string strConn;
			if (this.rdoFIADB.Checked==false)
			{
				strMDBFile = this.txtMDBPlot.Text.Trim();
				strTable = this.txtMDBPlotTable.Text.Trim();
				strConn = p_ado.getMDBConnString(strMDBFile,"","");
			}
			else
			{
				strMDBFile = this.txtMDBFiadbInputFile.Text.Trim();
				strTable = this.cmbFiadbPpsaTable.Text.Trim();
				strConn = p_ado.getMDBConnString(strMDBFile,"","");
			}

			if (this.chkNonForested.Checked == true && this.chkForested.Checked==true)
			{
				this.BuildFilterByStateCountyString("statecd","countycd",false);
				if (this.rdoFIADB.Checked==false)
				{
					p_ado.m_strSQL = "SELECT statecd,countycd,plot " + 
						"FROM " + strTable +  " " + 
						"WHERE " + this.m_strStateCountySQL.Trim() + " " + 
						"GROUP BY statecd,countycd,plot;";
				}
				else
				{
					p_ado.m_strSQL = "SELECT statecd,countycd,plot " + 
						"FROM " + strTable +  " " + 
						"WHERE RSCD = " + this.m_strCurrFIADBRsCd + " AND " + 
						"EVALID = " + this.m_strCurrFIADBEvalId + " AND " + this.m_strStateCountySQL.Trim() + " " + 
						"GROUP BY statecd,countycd,plot;";
				}
			}
			else if (this.chkForested.Checked==true)
			{
				if (this.rdoFIADB.Checked==false)
				{
					this.BuildFilterByStateCountyString("statecd","countycd",false);
					p_ado.m_strSQL = "SELECT statecd,countycd,plot " + 
						"FROM " + strTable + " " + 
						"WHERE plot_status_cd = 1 AND " + this.m_strStateCountySQL.Trim() + " " + 
						"GROUP BY statecd,countycd,plot;";
				}
				else
				{
					this.BuildFilterByStateCountyString("ppsa.statecd","ppsa.countycd",false);
					p_ado.m_strSQL = "SELECT ppsa.statecd,ppsa.countycd,ppsa.plot " + 
						"FROM " + strTable +  " ppsa " + 
						"INNER JOIN " + this.cmbFiadbPlotTable.Text.Trim() + " p " + 
						"ON ppsa.plt_cn = p.cn " + 
						"WHERE ppsa.RSCD = " + this.m_strCurrFIADBRsCd + " AND " + 
						"ppsa.EVALID = " + this.m_strCurrFIADBEvalId + " AND " +  
						"p.plot_status_cd=1 AND " + this.m_strStateCountySQL.Trim() + " " +
						"GROUP BY ppsa.statecd,ppsa.countycd,ppsa.plot;";
				}
			}
			else if (this.chkNonForested.Checked==true)
			{
				if (this.rdoFIADB.Checked==false)
				{
					this.BuildFilterByStateCountyString("statecd","countycd",false);
					p_ado.m_strSQL = "SELECT statecd,countycd,plot " + 
						"FROM " + strPlot + " " + 
						"WHERE plot_status_cd <> 1 AND " + this.m_strStateCountySQL.Trim() + " " +
						"GROUP BY statecd,countycd,plot;";
				}
				else
				{
					this.BuildFilterByStateCountyString("ppsa.statecd","ppsa.countycd",false);
					p_ado.m_strSQL = "SELECT ppsa.statecd,ppsa.countycd,ppsa.plot " + 
						"FROM " + strTable +  " ppsa " + 
						"INNER JOIN " + this.cmbFiadbPlotTable.Text.Trim() + " p " + 
						"ON ppsa.plt_cn = p.cn " + 
						"WHERE ppsa.RSCD = " + this.m_strCurrFIADBRsCd + " AND " + 
						"ppsa.EVALID = " + this.m_strCurrFIADBEvalId + " AND " +  
						"p.plot_status_cd<>1 AND " + this.m_strStateCountySQL.Trim() + " " +
						"GROUP BY ppsa.statecd,ppsa.countycd,ppsa.plot;";
				}
			}
			else
			{
				this.m_intError=-1;

			}

            
			if (p_ado.m_intError==0 && this.m_intError==0)
			{
				p_ado.SqlQueryReader(strConn,p_ado.m_strSQL);
				if (p_ado.m_intError==0)
				{
					try
					{
						//load up each row in the FIADB plot input table

						while (p_ado.m_OleDbDataReader.Read())
						{
							strState="";
							strCounty="";
							strPlot="";
							
							//make sure the row is not null values
							if (p_ado.m_OleDbDataReader[0] != System.DBNull.Value &&
								p_ado.m_OleDbDataReader[0].ToString().Trim().Length > 0)
							{
								strState= p_ado.m_OleDbDataReader["statecd"].ToString();
								strCounty = p_ado.m_OleDbDataReader["countycd"].ToString();
								strPlot = p_ado.m_OleDbDataReader["plot"].ToString();
								this.lstFilterByPlot.BeginUpdate();
								System.Windows.Forms.ListViewItem listItem = new ListViewItem();
								listItem.Checked=false;
								listItem.SubItems.Add(strState);
								listItem.SubItems.Add(strCounty);
								listItem.SubItems.Add(strPlot);
								this.lstFilterByPlot.Items.Add(listItem);
								this.lstFilterByPlot.EndUpdate();
								intAddedPlotRows++;
							}

						}
						p_ado.m_OleDbDataReader.Close();
						if (intAddedPlotRows == 0 )
						{
							this.m_intError=-1;
							MessageBox.Show("!!No Plots Loaded To Get State, County, Plot Information!!","Load State, County, Plot Menus");
						}
						((frmDialog)this.ParentForm).Enabled=true;
					}
					catch (Exception caught)
					{
						this.m_intError=-1;
						MessageBox.Show(caught.Message);
					}
					p_ado.m_OleDbConnection.Close();
				}
				else
				{
					this.m_intError=p_ado.m_intError;
				}
				
			}
			p_ado=null;
		}


		private void btnFilterByStatePrevious_Click(object sender, System.EventArgs e)
		{
			if (this.rdoFIADB.Checked==true && (this.rdoText.Checked==true || this.rdoAccess.Checked==true))
			{
				this.grpboxFilter.Visible=true;
				this.grpboxFilterByState.Visible=false;
			}
			else if (this.rdoIDB.Checked==true)
			{
				this.grpboxIDBInv.Visible=true;
				this.grpboxFilterByState.Visible=false;
			}
		}

		private void btnFilterByStateCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}

		private void btnFilterByStateNext_Click(object sender, System.EventArgs e)
		{
			if (this.rdoFIADB.Checked==true && this.rdoText.Checked==true)
			{
				this.txtFileInputPlot();
				if (this.m_intError == 0)
				{
					this.grpboxFilterByPlot.Visible=true;
					this.grpboxFilterByState.Visible=false;
				}
				
			}
			else if ((this.rdoFIADB.Checked==true && this.rdoAccess.Checked==true) || this.rdoIDB.Checked==true)
			{
				this.mdbInputPlot();
				if (this.m_intError==0)
				{
					this.grpboxFilterByPlot.Visible=true;
					this.grpboxFilterByState.Visible=false;
				}
			}   
		}

		private void btnFilterByPlotPrevious_Click(object sender, System.EventArgs e)
		{
			if (this.rdoFIADB.Checked==true && (this.rdoText.Checked==true || this.rdoAccess.Checked==true))
			{
				this.grpboxFilterByPlot.Visible = false;
				this.grpboxFilterByState.Visible=true;

			}
			else if (this.rdoIDB.Checked==true)
			{
				this.grpboxFilterByPlot.Visible = false;
				this.grpboxFilterByState.Visible=true;
			}
		}

		private void btnFilterByPlotCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();		
		}

		private void btnFilterByStateSelect_Click(object sender, System.EventArgs e)
		{
			for (int x=0; x<= this.lstFilterByState.Items.Count-1;x++)
			{
				this.lstFilterByState.Items[x].Checked=true;
			}
			
		}

		private void btnFilterByStateUnselect_Click(object sender, System.EventArgs e)
		{
			for (int x=0; x<= this.lstFilterByState.Items.Count-1;x++)
			{
				this.lstFilterByState.Items[x].Checked=false;
			}

		}

		private void btnFilterByPlotSelect_Click(object sender, System.EventArgs e)
		{
			for (int x=0; x<= this.lstFilterByPlot.Items.Count-1;x++)
			{
				this.lstFilterByPlot.Items[x].Checked=true;
			}

		}

		private void btnFilterByPlotUnselect_Click(object sender, System.EventArgs e)
		{
			for (int x=0; x<= this.lstFilterByPlot.Items.Count-1;x++)
			{
				this.lstFilterByPlot.Items[x].Checked=false;
			}

		}

		private void btnFilterByStateFinish_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).MinimizeMainForm=true;
			this.Enabled=false;
			this.m_strStateCountyPlotSQL="";
			this.m_strStateCountySQL="";
            this.m_intError=0;
			if (this.lstFilterByState.CheckedItems.Count==0) 
			{
				MessageBox.Show("Select At Least One State, County Item","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.Enabled = true;
				return;
			}
			if (this.rdoFIADB.Checked==true && this.rdoText.Checked==true)
			{
				if (this.chkForested.Checked && this.chkNonForested.Checked)
					this.BuildFilterByStateCountyString("statecd","countycd",true);
				else
					this.BuildFilterByStateCountyString("ppsa.statecd","ppsa.countycd",true);

				if (this.m_intError==0)
				{

                    LoadTxtPlotCondTreeData_Start();
				}
			}
			else if (this.rdoFIADB.Checked==true && this.rdoAccess.Checked==true)
			{
                CalculateAdjustments_Start();
                if (m_intError == 0)
                {
                    if (this.chkForested.Checked && this.chkNonForested.Checked)
                        this.BuildFilterByStateCountyString("statecd", "countycd", true);
                    else
                        this.BuildFilterByStateCountyString("ppsa.statecd", "ppsa.countycd", true);
                    if (this.m_intError == 0)
                    {

                        this.LoadMDBPlotCondTreeData_Start();
                    }
                }

			}
			else if (this.rdoIDB.Checked==true)
			{
				this.BuildFilterByStateCountyString("statecd","countycd",false);
				if (this.m_intError==0)
				{
                    this.LoadIDBPlotCondTreeData_Start();
				}
			}
           
		
		}
		private void BuildFilterByStateCountyString(string strStateFieldAlias,string strCountyFieldAlias,bool bStringDataType)
		{
			string strCurState="";
			string strState="";
			string strCounty="";
			bool bAllCounties;
			string strStateList="";
			string strCountyList="";
			int y=0;

            System.Windows.Forms.ListView oLv = 
                (System.Windows.Forms.ListView)frmMain.g_oDelegate.GetListView(lstFilterByState, false);

            int intTotalCount = (int)frmMain.g_oDelegate.GetListViewItemsPropertyValue(oLv, "Count", false);
            int intCheckedCount = (int)frmMain.g_oDelegate.GetListViewCheckedItemsCount(oLv, false);

			this.m_strStateCountyPlotSQL="";
			if (intTotalCount == intCheckedCount)
			{
				this.m_bAllCountiesSelected = true;
				bAllCounties=true;
			}
			else
			{
				this.m_bAllCountiesSelected = false;
				bAllCounties=false;
			}
			this.m_strStateCountySQL="";
			
			for (int x=0; x <= intTotalCount -1;x++)
			{
                if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, x, "Checked", false))
				{
                    strState = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, 1, "Text", false);
                    strCounty = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, 2, "Text", false);
                    strState = strState.Trim(); strCounty = strCounty.Trim();

					if (strState.Trim().Length > 0)
					{
					}
					//check to see if this is a new state
					if (strState !=
						strCurState && strState.Trim().Length > 0)

					{

						if (strCurState.Trim().Length == 0)
						{
							//first time
							strCurState = strState;
						}
						if (this.m_bAllCountiesSelected == true)
						{
							if (strStateList.Trim().Length ==0)
							{
								if (bStringDataType == false)
								{
									strStateList = strState;
								}
								else
								{
									strStateList = "'" + strState.Trim() + "'";
								}
							}
							else
							{
								if (bStringDataType == false)
								{
									strStateList += "," + strState;
								}
								else
								{
									strStateList += ",'" + strState.Trim() + "'";
								}
							}
							strCurState=strState;
						}
						else
						{
							
							//current state
							//check if all counties for this state are selected
							strCountyList="";
							bAllCounties=true;

							//check to see if the previous list item is the same state and if
							//it is checked
							if (x-1 >=0)
							{
								if ((string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv,x-1,1,"Text",false).ToString().Trim() == strState.Trim() && 
                                    (bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv,x-1,"Checked",false)==false)
								{
									bAllCounties=false;
								}
							}
							
							for (y=x;y<=intTotalCount-1;y++)
							{
                                if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, y, "Checked", false) == true)
								{

									if (strState.Trim() !=
                                        (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, y, 1, "Text", false).ToString().Trim())
									{
										break;
									}
                                    strCounty = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, y, 2, "Text", false).ToString().Trim();
									if (strCountyList.Trim().Length ==0)
									{
										if (bStringDataType == false)
										{
											strCountyList = strCounty;
										}
										else
										{
											strCountyList = "'" + strCounty.Trim() + "'";
										}
									}
									else
									{
										if (bStringDataType == false)
										{
											strCountyList += "," + strCounty;
										}
										else
										{
											strCountyList += ",'" + strCounty.Trim() + "'";
										}
									}

								}
								else
								{
                                    if (strState.Trim() == (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, y, 1, "Text", false).ToString().Trim())
									{
										bAllCounties=false;
									}
								}
							}
							strCurState=strState;
							if (y<=intTotalCount-1)
							{
								x = y - 1;
							}
							else
							{
								x = y;
							}
							if (bAllCounties==true)
							{
								if (this.m_strStateCountySQL.Trim().Length > 0)
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountySQL += " OR (" + strStateFieldAlias  + " = " + strCurState + ")";
									}
									else
									{
										this.m_strStateCountySQL += " OR ( trim(" + strStateFieldAlias.Trim()  + ") = '" + strCurState.Trim() + "')";
									}
								}
								else
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountySQL = "(" + strStateFieldAlias  + " = " + strCurState + ")";
									}
									else
									{
										this.m_strStateCountySQL = "( trim(" + strStateFieldAlias.Trim()  + ") = '" + strCurState.Trim() + "')";
									}
								}
							}
							else
							{
								if (this.m_strStateCountySQL.Trim().Length > 0)
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountySQL += " OR (" + strStateFieldAlias + " = " + strCurState  + " AND " + strCountyFieldAlias + " IN (" + strCountyList + "))";
									}
									else
									{
										this.m_strStateCountySQL += " OR ( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim()  + "' AND trim(" + strCountyFieldAlias.Trim() + ") IN (" + strCountyList + "))";
									}
								}
								else
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountySQL = "(" + strStateFieldAlias + " = " + strCurState  + " AND " + strCountyFieldAlias + " IN (" + strCountyList + "))";
									}
									else
									{
										this.m_strStateCountySQL = "( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim()  + "' AND trim(" + strCountyFieldAlias.Trim() + ") IN (" + strCountyList + "))";
									}
								}
							}
						}

					}
				}
			}
			if (this.m_bAllCountiesSelected==true)
			{
				if (bStringDataType==false)
				{
					if (rdoFIADB.Checked)
					{
						this.m_strStateCountySQL = "(" + strStateFieldAlias + " IN (" + strStateList + "))";
					}
					else
					{
						if (this.m_strIDBInv.Trim().Length == 0)
						{
							this.m_strStateCountySQL = "(" + strStateFieldAlias + " IN (" + strStateList + "))";
						}
						else
						{
							this.m_strStateCountySQL = "(MID(biosum_plot_id,2,4)='" + 
								this.m_strIDBInv.Trim() + "' AND " + 
								strStateFieldAlias + " IN (" + strStateList + ") AND " + 
								"(plot NOT BETWEEN 99000 AND 99999))";
						}
					}
				}
				else
				{
					this.m_strStateCountySQL = "( trim(" + strStateFieldAlias.Trim() + ") IN (" + strStateList + "))";
				}
			}
			else
			{
				if (rdoFIADB.Checked)
				{
					this.m_strStateCountySQL = "(" + this.m_strStateCountySQL + ")";
				}
				else
				{
			
					if (this.m_strIDBInv.Trim().Length == 0)
					{
						this.m_strStateCountySQL = "(" + this.m_strStateCountySQL + ")";
					}
					else
					{
						this.m_strStateCountySQL = "(MID(biosum_plot_id,2,4)='" + 
							this.m_strIDBInv.Trim() + "' AND " + this.m_strStateCountySQL + " AND " + 
							"(plot NOT BETWEEN 99000 AND 99999))";

					}
				}
				
			}
		}
		private void BuildFilterByPlotString(string strStateFieldAlias,string strCountyFieldAlias,string strPlotFieldAlias, bool bStringDataType)
		{
			string strCurState="";
			string strCurCounty="";
			string strState="";
			string strCounty="";
			string strPlot="";
			bool bAllPlots;
			string strCountyList="";
			string strPlotList = "";
			int y=0;
         
			this.m_strStateCountySQL="";
            System.Windows.Forms.ListView oLv;

            oLv = (System.Windows.Forms.ListView)frmMain.g_oDelegate.GetListView(lstFilterByPlot, false);
            int intPlotCount = (int)frmMain.g_oDelegate.GetListViewItemsPropertyValue(oLv, "Count", false);

            if ((int)frmMain.g_oDelegate.GetListViewCheckedItemsCount(oLv, false) == 
                intPlotCount)
			{
				this.m_bAllPlotsSelected = true;
				bAllPlots=true;
			}
			else
			{
				this.m_bAllPlotsSelected = false;
				bAllPlots=false;
			}
			this.m_strStateCountyPlotSQL="";
			
			for (int x=0; x <= intPlotCount -1;x++)
			{
				if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv,x,"Checked",false))
				{
					strState = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv,x,1,"Text",false);
                    strCounty = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, 2, "Text", false);
                    strPlot = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, 3, "Text", false);
                    strState = strState.Trim(); strCounty = strCounty.Trim(); strPlot = strPlot.Trim();
					if (strState.Trim().Length > 0)
					{
					}
					//check to see if this is a new state
					if ((strState !=
						strCurState && strState.Trim().Length > 0) ||
						(strCounty != strCurCounty && strCounty.Trim().Length > 0))

					{

						if (strCurState.Trim().Length == 0)
						{
							//first time
							strCurState = strState;
							strCurCounty = strCounty;
							strCountyList="";

						}
						if (this.m_bAllPlotsSelected == true)
						{
							if (strCurState.Trim() != strState.Trim())
							{
								if (this.m_strStateCountyPlotSQL.Trim().Length >0)
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountyPlotSQL += " OR (" + strStateFieldAlias + " = " + strCurState + " AND " + strCountyFieldAlias + " IN (" + strCountyList + "))";
									}
									else
									{
										this.m_strStateCountyPlotSQL += " OR ( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim() + "' AND trim(" + strCountyFieldAlias.Trim() + ") IN (" + strCountyList + "))";
									}

								}
								else
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountyPlotSQL = "(" + strStateFieldAlias + " = " + strCurState + " AND " + strCountyFieldAlias + " IN (" + strCountyList + "))";
									}
									else
									{
										this.m_strStateCountyPlotSQL = "( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim() + "' AND trim(" + strCountyFieldAlias.Trim() + ") IN (" + strCountyList + "))";
									}

								}
								strCountyList="";

							}

							if (strCountyList.Trim().Length ==0)
							{
								if (bStringDataType == false)
								{
									strCountyList = strCounty;
								}
								else
								{
									strCountyList = "'" + strCounty.Trim() + "'";
								}
							}
							else
							{
								if (bStringDataType == false)
								{
									strCountyList += "," + strCounty;
								}
								else
								{
									strCountyList += ",'" + strCounty.Trim() + "'";
								}
							}
							strCurState=strState;
							strCurCounty=strCounty;
						}
						else
						{
							
							//current state and county
							//check if all plots for this state and county are selected
							strCountyList="";
							strPlotList="";
							bAllPlots=true;

							//check to see if the previous list item is the same state and if
							//it is checked
							if (x-1 >=0)
							{
								if ((string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv,x-1,1,"Text",false).ToString().Trim() == strState.Trim() && 
                                    (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv,x-1,2,"Text",false).ToString().Trim() == strCounty.Trim() && 
									(bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv,x-1,"Checked",false)==false)
								{
									bAllPlots=false;
								}
							}
							
							for (y=x;y<= intPlotCount-1;y++)
							{
                                if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, y, "Checked", false) == true)
								{

									if (strState.Trim() != 
                                          (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv,y,1,"Text",false).ToString().Trim() ||
										strCounty.Trim() != (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv,y,2,"Text",false).ToString().Trim())
									{
										break;
									}
                                    strPlot = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, y, 3, "Text", false).ToString().Trim();
									if (strPlotList.Trim().Length ==0)
									{
										if (bStringDataType == false)
										{
											strPlotList = strPlot;
										}
										else
										{
											strPlotList = "'" + strPlot.Trim() + "'";
										}
									}
									else
									{
										if (bStringDataType == false)
										{
											strPlotList += "," + strPlot;
										}
										else
										{
											strPlotList += ",'" + strPlot.Trim() + "'";
										}
									}

								}
								else
								{
									if (strState.Trim() == (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv,y,1,"Text",false).ToString().Trim() && 
										strCounty.Trim() == (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv,y,2,"Text",false).ToString().Trim())
									{
										bAllPlots=false;
									}
								}
							}
							strCurState=strState;
							strCurCounty = strCounty;
							if (y<=intPlotCount-1)
							{
								x = y - 1;
							}
							else
							{
								x = y;
							}
							if (bAllPlots==true)
							{
								if (this.m_strStateCountyPlotSQL.Trim().Length > 0)
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountyPlotSQL += " OR (" + strStateFieldAlias  + " = " + strCurState + " AND " + strCountyFieldAlias + " = " + strCurCounty + ")";
									}
									else
									{
										this.m_strStateCountyPlotSQL += " OR ( trim(" + strStateFieldAlias.Trim()  + ") = '" + strCurState.Trim() + "'  AND trim(" + strCountyFieldAlias.Trim() + ") = '" + strCurCounty.Trim() + "')";
									}
								}
								else
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountyPlotSQL = "(" + strStateFieldAlias  + " = " + strCurState + " AND " + strCountyFieldAlias + " = " + strCurCounty + ")";
									}
									else
									{
										this.m_strStateCountyPlotSQL = "( trim(" + strStateFieldAlias.Trim()  + ") = '" + strCurState.Trim() + "'  AND trim(" + strCountyFieldAlias.Trim() + ") = '" + strCurCounty.Trim() + "')";
									}
								}
							}
							else
							{
								if (this.m_strStateCountyPlotSQL.Trim().Length > 0)
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountyPlotSQL += " OR (" + strStateFieldAlias + " = " + strCurState  + " AND " + strCountyFieldAlias + " = " + strCurCounty + " AND " + strPlotFieldAlias + " IN (" + strPlotList + "))";
									}
									else
									{
										this.m_strStateCountyPlotSQL += " OR ( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim()  + "' AND trim(" + strCountyFieldAlias.Trim() + ") = '" + strCurCounty.Trim() + "' AND trim(" + strPlotFieldAlias.Trim() + ") IN (" + strPlotList + "))";
									}
								}
								else
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountyPlotSQL = "(" + strStateFieldAlias + " = " + strCurState  + " AND " + strCountyFieldAlias + " = " + strCurCounty + " AND " + strPlotFieldAlias + " IN (" + strPlotList + "))";
									}
									else
									{
										this.m_strStateCountyPlotSQL = "( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim()  + "' AND trim(" + strCountyFieldAlias.Trim() + ") = '" + strCurCounty.Trim() + "' AND trim(" + strPlotFieldAlias.Trim() + ") IN (" + strPlotList + "))";
									}
								}
							}
						}
					}
				}
			}
			if (this.m_bAllPlotsSelected == true)
			{
				
				if (this.m_strStateCountyPlotSQL.Trim().Length >0)
				{
					if (bStringDataType == false)
					{
						this.m_strStateCountyPlotSQL += " OR (" + strStateFieldAlias + " = " + strCurState + " AND " + strCountyFieldAlias + " IN (" + strCountyList + "))";
					}
					else
					{
						this.m_strStateCountyPlotSQL += " OR ( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim() + "' AND trim(" + strCountyFieldAlias.Trim() + ") IN (" + strCountyList + "))";
					}

				}
				else
				{
					if (bStringDataType == false)
					{
						this.m_strStateCountyPlotSQL = "(" + strStateFieldAlias + " = " + strCurState + " AND " + strCountyFieldAlias + " IN (" + strCountyList + "))";
					}
					else
					{
						this.m_strStateCountyPlotSQL = "( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim() + "' AND trim(" + strCountyFieldAlias.Trim() + ") IN (" + strCountyList + "))";
					}

				}
			}
			this.m_strStateCountyPlotSQL = "(" + this.m_strStateCountyPlotSQL + ")";
			//MessageBox.Show(this.m_strStateCountyPlotSQL);
		
	
			
		}

		private void btnFilterByPlotFinish_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).MinimizeMainForm=true;
			this.Enabled=false;
            m_intError = 0;
			this.m_strStateCountyPlotSQL="";
			this.m_strStateCountySQL="";

			if (this.lstFilterByPlot.CheckedItems.Count==0) 
			{
				MessageBox.Show("Select At Least One State, County, Plot Item","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.Enabled = true;
				return;
			}
			if (this.rdoFIADB.Checked==true && this.rdoText.Checked==true)
			{

				this.BuildFilterByPlotString("ppsa.statecd","ppsa.countycd","ppsa.plot",true);
				if (this.m_intError==0)
				{

                    LoadTxtPlotCondTreeData_Start();
				}
			}
			else if (this.rdoFIADB.Checked==true && this.rdoAccess.Checked==true)
			{
                CalculateAdjustments_Start();
                if (m_intError == 0)
                {
                    this.BuildFilterByPlotString("ppsa.statecd", "ppsa.countycd", "ppsa.plot", false);
                    if (this.m_intError == 0)
                    {

                        this.LoadMDBPlotCondTreeData_Start();
                    }
                }
			}
			else if (this.rdoIDB.Checked==true)
			{
				this.BuildFilterByPlotString("statecd","countycd","plot",false);
				if (this.m_intError==0)
				{
                    this.LoadIDBPlotCondTreeData_Start();
				}
			}
			//((frmDialog)this.ParentForm).MinimizeMainForm=false;
			//this.Enabled=true;
		
		}

		private void btnMDBInputCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}

		private void btnMDBInputPrevious_Click(object sender, System.EventArgs e)
		{
			this.grpboxMDBInput.Visible=false;
			this.grpboxInvType.Visible=true;
		}

		private void btnMDBPlotBrowse_Click(object sender, System.EventArgs e)
		{
			this.GetMDBFileAndTable("MS Access Data File Containing Plot Table Data",
				"Select Plot Table",
				ref this.txtMDBPlot,
				ref this.txtMDBPlotTable);
		}

		private void btnMDBCondBrowse_Click(object sender, System.EventArgs e)
		{
			this.GetMDBFileAndTable("MS Access Data File Containing Condition Table Data",
				"Select Condition Table",
				ref this.txtMDBCond,
				ref this.txtMDBCondTable);
		}

		private void btnMDBTreeBrowse_Click(object sender, System.EventArgs e)
		{
			this.GetMDBFileAndTable("MS Access Data File Containing Tree Table Data",
				"Select Tree Table",
				ref this.txtMDBTree,
				ref this.txtMDBTreeTable);
		}

		private void btnMDBInputNext_Click(object sender, System.EventArgs e)
		{
			if (this.txtMDBPlot.Text.Trim().Length == 0)
			{
				MessageBox.Show("Select A Plot MDB File And Table","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				return;
			}
			if (this.txtMDBCond.Text.Trim().Length == 0)
			{
				MessageBox.Show("Select A Cond MDB File And Table","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				return;
			}
			if (this.txtMDBTree.Text.Trim().Length == 0)
			{
				MessageBox.Show("Select A Tree MDB File And Table","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				return;
			}
			if (this.rdoFIADB.Checked==true)
			{
				if (this.rdoFilterNone.Checked==true)
				{
					this.btnFilterNext.Enabled=false;
					this.btnFilterFinish.Enabled=true;
				}
				else if (this.rdoFilterByMenu.Checked==true)
				{
					this.btnFilterNext.Enabled=true;
					this.btnFilterFinish.Enabled=false;
                  
				}
				else
				{
					this.btnFilterNext.Enabled=false;
					this.btnFilterFinish.Enabled=true;
				}
				this.rdoFilterByFile.Text = "Filter By File (Text File Containing Plot_CN numbers)";
			}
			else
			{
				if (this.rdoFilterByFile.Checked==true) 
				{
					this.btnFilterNext.Enabled=false;
					this.btnFilterFinish.Enabled=true;
				}
				else if (this.rdoFilterByMenu.Checked==true)
				{
					this.btnFilterNext.Enabled=true;
					this.btnFilterFinish.Enabled=false;
                  
				}
				else
				{
					this.btnFilterNext.Enabled=true;
					this.btnFilterFinish.Enabled=false;
				}
				this.rdoFilterByFile.Text = "Filter By File (Text File Containing idb_plot_id numbers)";
			}

			this.grpboxFilter.Visible=true;
			this.grpboxMDBInput.Visible=false;



		}

		private void btnIDBInvPrevious_Click(object sender, System.EventArgs e)
		{
			this.grpboxFilter.Visible=true;
			this.grpboxIDBInv.Visible=false;
		}

		private void btnIDBInvCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}
		private void idbLoadInv()
		{
            string strId="";
			string strDS="";
			string strDesc="";
			string strDef="";
			int intAddedRows=0;
    

			      
			this.lstIDBInv.Clear();
			this.lstIDBInv.Columns.Add("Id", 30, HorizontalAlignment.Left);
			this.lstIDBInv.Columns.Add("Data Source", 100, HorizontalAlignment.Left);
			this.lstIDBInv.Columns.Add("Description", 100, HorizontalAlignment.Left);
			this.lstIDBInv.Columns.Add("Definition", 100, HorizontalAlignment.Left);
            this.m_strIDBInv="";
			this.m_intError=0;

			

			FIA_Biosum_Manager.ado_data_access p_ado = new ado_data_access();

			
			//get all the project datasources

			string strMDBFile = ((frmDialog)this.ParentForm).m_frmMain.frmProject.uc_project1.txtRootDirectory.Text.ToString().Trim() + "\\db\\ref_master.mdb";
			string strConn = p_ado.getMDBConnString(strMDBFile,"","");
			p_ado.m_strSQL = "select * from inventories order by idb_data_source";

	

           
				p_ado.SqlQueryReader(strConn,p_ado.m_strSQL);
				if (p_ado.m_intError==0)
				{
					try
					{
						//load up each row in the FIADB plot input table

						while (p_ado.m_OleDbDataReader.Read())
						{
							strDS="";
							strDesc="";
							strDef="";
							
							//make sure the row is not null values
							if (p_ado.m_OleDbDataReader[0] != System.DBNull.Value &&
								p_ado.m_OleDbDataReader[0].ToString().Trim().Length > 0)
							{
								strId =p_ado.m_OleDbDataReader["inv_id"].ToString();
								strDS= p_ado.m_OleDbDataReader["idb_data_source"].ToString();
								strDesc = p_ado.m_OleDbDataReader["description"].ToString();
								strDef = p_ado.m_OleDbDataReader["inv_id_def"].ToString();
								this.lstIDBInv.BeginUpdate();
								System.Windows.Forms.ListViewItem listItem = new ListViewItem();
								listItem.Text=strId;
								listItem.SubItems.Add(strDS);
								listItem.SubItems.Add(strDesc);
								listItem.SubItems.Add(strDef);
								this.lstIDBInv.Items.Add(listItem);
								this.lstIDBInv.EndUpdate();
								intAddedRows++;
							}

						}
						p_ado.m_OleDbDataReader.Close();
						if (intAddedRows == 0 )
						{
							this.m_intError=-1;
							MessageBox.Show("!!No Inventories Loaded !","Load Inventories");
						}
						else
						{
							this.lstIDBInv.Columns[0].Width = -1;
							this.lstIDBInv.Columns[1].Width = -1;
							this.lstIDBInv.Columns[2].Width = -1;
							this.lstIDBInv.Columns[3].Width = -1;
						}
						((frmDialog)this.ParentForm).Enabled=true;
					}
					catch (Exception caught)
					{
						this.m_intError=-1;
						MessageBox.Show(caught.Message);
					}
					p_ado.m_OleDbConnection.Close();
				}
				else
				{
					this.m_intError=p_ado.m_intError;
				}
			
			
			
			p_ado=null;
		}

		private void FIADBLoadInv()
		{
			string strEvalId="";
			string strRsCd="";
			string strStateCd="";
			string strLocNm="";
			string strEvalDesc="";
			string strRptYr="";
			string strNotes="";
			int intAddedRows=0;
			
			      
			this.lstFIADBInv.Clear();
			this.lstFIADBInv.Columns.Add("EvalId", 50, HorizontalAlignment.Left);
			this.lstFIADBInv.Columns.Add("RsCd", 50, HorizontalAlignment.Left);
			this.lstFIADBInv.Columns.Add("StateCd", 50, HorizontalAlignment.Left);
			this.lstFIADBInv.Columns.Add("Location_Nm", 100, HorizontalAlignment.Left);
			this.lstFIADBInv.Columns.Add("Eval_Descr", 300, HorizontalAlignment.Left);
			this.lstFIADBInv.Columns.Add("ReportYear", 200, HorizontalAlignment.Left);
			this.lstFIADBInv.Columns.Add("Notes", 200, HorizontalAlignment.Left);
//			this.m_strFIADBEvalId="";
//			this.m_strFIADBRsCd="";
			this.m_intError=0;

			

			FIA_Biosum_Manager.ado_data_access p_ado = new ado_data_access();

			
			//get all the project datasources


			string strMDBFile = ((frmDialog)this.ParentForm).m_frmMain.frmProject.uc_project1.txtRootDirectory.Text.ToString().Trim() + "\\db\\master.mdb";
			string strConn = p_ado.getMDBConnString(strMDBFile,"","");

			p_ado.m_strSQL = "SELECT * FROM " + this.m_strPopEvalTable + " where biosum_status_cd=9 order by statecd,evalid";

	

           
			p_ado.SqlQueryReader(strConn,p_ado.m_strSQL);
			if (p_ado.m_intError==0)
			{
				try
				{
					//load up each row in the FIADB plot input table

					while (p_ado.m_OleDbDataReader.Read())
					{
						strEvalId="";
						strRsCd="";
						strStateCd="";
						strLocNm="";
						strEvalDesc="";
						strRptYr="";
						strNotes="";
							
						//make sure the row is not null values
						if (p_ado.m_OleDbDataReader[0] != System.DBNull.Value &&
							p_ado.m_OleDbDataReader[0].ToString().Trim().Length > 0)
						{
							strEvalId =p_ado.m_OleDbDataReader["evalid"].ToString();
							strRsCd= p_ado.m_OleDbDataReader["RsCd"].ToString();
							strStateCd = p_ado.m_OleDbDataReader["statecd"].ToString();
							if (p_ado.m_OleDbDataReader["location_nm"] != System.DBNull.Value)
								strLocNm = p_ado.m_OleDbDataReader["location_nm"].ToString();
							if (p_ado.m_OleDbDataReader["eval_descr"] != System.DBNull.Value)
								strEvalDesc = p_ado.m_OleDbDataReader["eval_descr"].ToString();
							if (p_ado.m_OleDbDataReader["report_year_nm"] != System.DBNull.Value)
								strRptYr = p_ado.m_OleDbDataReader["report_year_nm"].ToString();
							if (p_ado.m_OleDbDataReader["notes"] != System.DBNull.Value)
								strNotes = p_ado.m_OleDbDataReader["notes"].ToString();

							this.lstFIADBInv.BeginUpdate();
							System.Windows.Forms.ListViewItem listItem = new ListViewItem();
							listItem.Text=strEvalId;
							listItem.SubItems.Add(strRsCd);
							listItem.SubItems.Add(strStateCd);
							listItem.SubItems.Add(strLocNm);
							listItem.SubItems.Add(strEvalDesc);
							listItem.SubItems.Add(strRptYr);
							listItem.SubItems.Add(strNotes);
							this.lstFIADBInv.Items.Add(listItem);
							this.lstFIADBInv.EndUpdate();
							intAddedRows++;
						}

					}
					p_ado.m_OleDbDataReader.Close();
					p_ado.m_OleDbDataReader=null;
					if (intAddedRows == 0 )
					{
						this.m_intError=-1;
						MessageBox.Show("!!No Inventories Loaded !","Load Inventories");
					}
					else
					{
					}
					((frmDialog)this.ParentForm).Enabled=true;
				}
				catch (Exception caught)
				{
					this.m_intError=-1;
					MessageBox.Show(caught.Message);
				}
				p_ado.m_OleDbConnection.Close();
				while (p_ado.m_OleDbConnection.State == System.Data.ConnectionState.Open)
					System.Threading.Thread.Sleep(1000);
				p_ado.m_OleDbConnection.Dispose();
				p_ado.m_OleDbConnection=null;


			}
			else
			{
				this.m_intError=p_ado.m_intError;
			}
			p_ado=null;
		}


		private void btnIDBInvSelectAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0; x<= this.lstIDBInv.Items.Count-1;x++)
			{
				this.lstIDBInv.Items[x].Checked=true;
			}
		}

		private void btnIDBInvClearAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0; x<= this.lstIDBInv.Items.Count-1;x++)
			{
				this.lstIDBInv.Items[x].Checked=false;
			}
		}


		private void BuildFilterByIDBInvString(string strStateFieldAlias,string strCountyFieldAlias,bool bStringDataType)
		{
		
			//string strCurInv;
			string strCurState="";
			//string strInv="";
			string strState="";
			string strCounty="";
			bool bAllCounties;
			string strStateList="";
			string strCountyList="";
			int y=0;
         
			this.m_strIDBInv="";
			if (this.lstFilterByState.CheckedItems.Count == this.lstFilterByState.Items.Count)
			{
				this.m_bAllCountiesSelected = true;
				bAllCounties=true;
			}
			else
			{
				this.m_bAllCountiesSelected = false;
				bAllCounties=false;
			}
			this.m_strStateCountySQL="";
			
			for (int x=0; x <= this.lstFilterByState.Items.Count -1;x++)
			{
				if (this.lstFilterByState.Items[x].Checked==true)
				{
					strState = this.lstFilterByState.Items[x].SubItems[1].Text.Trim();
					strCounty = this.lstFilterByState.Items[x].SubItems[2].Text.Trim();
					if (strState.Trim().Length > 0)
					{
					}
					//check to see if this is a new state
					if (strState !=
						strCurState && strState.Trim().Length > 0)

					{

						if (strCurState.Trim().Length == 0)
						{
							//first time
							strCurState = strState;
						}
						if (this.m_bAllCountiesSelected == true)
						{
							if (strStateList.Trim().Length ==0)
							{
								if (bStringDataType == false)
								{
									strStateList = strState;
								}
								else
								{
									strStateList = "'" + strState.Trim() + "'";
								}
							}
							else
							{
								if (bStringDataType == false)
								{
									strStateList += "," + strState;
								}
								else
								{
									strStateList += ",'" + strState.Trim() + "'";
								}
							}
							strCurState=strState;
						}
						else
						{
							
							//current state
							//check if all counties for this state are selected
							strCountyList="";
							bAllCounties=true;

							//check to see if the previous list item is the same state and if
							//it is checked
							if (x-1 >=0)
							{
								if (this.lstFilterByState.Items[x-1].SubItems[1].Text.Trim() == 
									strState.Trim() && this.lstFilterByState.Items[x-1].Checked==false)
								{
									bAllCounties=false;
								}
							}
							
							for (y=x;y<=this.lstFilterByState.Items.Count-1;y++)
							{
								if (this.lstFilterByState.Items[y].Checked==true)
								{

									if (strState.Trim() != this.lstFilterByState.Items[y].SubItems[1].Text.Trim())
									{
										break;
									}
									strCounty =  this.lstFilterByState.Items[y].SubItems[2].Text.Trim();
									if (strCountyList.Trim().Length ==0)
									{
										if (bStringDataType == false)
										{
											strCountyList = strCounty;
										}
										else
										{
											strCountyList = "'" + strCounty.Trim() + "'";
										}
									}
									else
									{
										if (bStringDataType == false)
										{
											strCountyList += "," + strCounty;
										}
										else
										{
											strCountyList += ",'" + strCounty.Trim() + "'";
										}
									}

								}
								else
								{
									if (strState.Trim() == this.lstFilterByState.Items[y].SubItems[1].Text.Trim())
									{
										bAllCounties=false;
									}
								}
							}
							strCurState=strState;
							if (y<=this.lstFilterByState.Items.Count-1)
							{
								x = y - 1;
							}
							else
							{
								x = y;
							}
							if (bAllCounties==true)
							{
								if (this.m_strStateCountySQL.Trim().Length > 0)
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountySQL += " OR (" + strStateFieldAlias  + " = " + strCurState + ")";
									}
									else
									{
										this.m_strStateCountySQL += " OR ( trim(" + strStateFieldAlias.Trim()  + ") = '" + strCurState.Trim() + "')";
									}
								}
								else
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountySQL = "(" + strStateFieldAlias  + " = " + strCurState + ")";
									}
									else
									{
										this.m_strStateCountySQL = "( trim(" + strStateFieldAlias.Trim()  + ") = '" + strCurState.Trim() + "')";
									}
								}
							}
							else
							{
								if (this.m_strStateCountySQL.Trim().Length > 0)
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountySQL += " OR (" + strStateFieldAlias + " = " + strCurState  + " AND " + strCountyFieldAlias + " IN (" + strCountyList + "))";
									}
									else
									{
										this.m_strStateCountySQL += " OR ( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim()  + "' AND trim(" + strCountyFieldAlias.Trim() + ") IN (" + strCountyList + "))";
									}
								}
								else
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountySQL = "(" + strStateFieldAlias + " = " + strCurState  + " AND " + strCountyFieldAlias + " IN (" + strCountyList + "))";
									}
									else
									{
										this.m_strStateCountySQL = "( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim()  + "' AND trim(" + strCountyFieldAlias.Trim() + ") IN (" + strCountyList + "))";
									}
								}
							}
						}

					}
				}
			}
			if (this.m_bAllCountiesSelected==true)
			{
				if (bStringDataType==false)
				{
					this.m_strStateCountySQL = "(" + strStateFieldAlias + " IN (" + strStateList + "))";
				}
				else
				{
					this.m_strStateCountySQL = "( trim(" + strStateFieldAlias.Trim() + ") IN (" + strStateList + "))";
				}
			}
			else
			{
				this.m_strStateCountySQL = "(" + this.m_strStateCountySQL + ")";
			}
		}

		private void lstIDBInv_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
		{
			
		}

		private void btnIDBInvNext_Click(object sender, System.EventArgs e)
		{
			if (this.lstIDBInv.SelectedItems.Count > 0)
			{
				this.m_strIDBInv = this.lstIDBInv.SelectedItems[0].Text.Trim();
				this.mdbInputStateCounty();
				this.grpboxIDBInv.Visible=false;
				this.grpboxFilterByState.Visible=true;
			}
			else
			{
				MessageBox.Show("Select A PNW IDB Inventory","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			}
		}

		private void btnIDBInvAppend_Click(object sender, System.EventArgs e)
		{
			if (this.lstIDBInv.SelectedItems.Count > 0)
			{
				
				
				this.m_strIDBInv = this.lstIDBInv.SelectedItems[0].Text.Trim();
                this.LoadIDBPlotCondTreeData_Start();
			}
			else
			{
				MessageBox.Show("Select A PNW IDB Inventory","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			}

		
		}
        private void LoadIDBPlotCondTreeData_Process()
		{
            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
			string strFields="";
			string strValues="";
			int intAddedPlotRows=0;
			int intAddedCondRows=0;
			int intAddedTreeRows=0;
			int intAddedSiteTreeRows=0;
			int x=0;
			int y=0;
			string strCol="";
			string strTime="";
            string str = "";
            string str2 = "";



			this.m_intError=0;	
						
			
            
			try
			{
				//-----------PREPARATION FOR ADDING PLOT RECORDS---------//
                    
				//instatiate the oledb data access class
				this.m_ado = new ado_data_access();


                //progress bar 1: single process
                this.SetThermValue(m_frmTherm.progressBar1, "Maximum", 100);
                this.SetThermValue(m_frmTherm.progressBar1, "Minimum", 0);
                this.SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                this.SetLabelValue(m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);

                //progress bar 2: overall progress
                this.SetThermValue(m_frmTherm.progressBar2, "Maximum", 100);
                this.SetThermValue(m_frmTherm.progressBar2, "Minimum", 0);
                this.SetThermValue(m_frmTherm.progressBar2, "Value", 0);
                this.SetLabelValue(m_frmTherm.lblMsg2, "Text", "Overall Progress");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);




				//get all the project datasources
                FIA_Biosum_Manager.Datasource p_datasource = new Datasource();
                p_datasource.m_strDataSourceMDBFile = this.ReferenceFormDialog.m_frmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\project.mdb";
                p_datasource.LoadTableColumnNamesAndDataTypes = false;
                p_datasource.LoadTableRecordCount = false;
                p_datasource.m_strDataSourceTableName = "datasource";
                p_datasource.m_strScenarioId = "";
                p_datasource.populate_datasource_array();

				//create a temporary mdb file with links to all the project tables
				//and return the name of the file that contains the links
				this.m_strTempMDBFile = p_datasource.CreateMDBAndTableDataSourceLinks();

			
				//instatiate dao for creating links in the temp table
				//to the fiadb plot, cond, and tree input tables
				dao_data_access p_dao1 = new dao_data_access();
                this.SetLabelValue(m_frmTherm.lblMsg, "Text", "Creating Datasource Links");

				//create links to the idb input tables in the temp mdb file
                //create links to the fiadb input tables in the temp mdb file
                //plot table
                str = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBPlot, "Text", false);
                str2 = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBPlotTable, "Text", false);
                p_dao1.CreateTableLink(this.m_strTempMDBFile, "idb_plot_input", str.Trim(), str2.Trim());
                //cond table
                str = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBCond, "Text", false);
                str2 = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBCondTable, "Text", false);
                p_dao1.CreateTableLink(this.m_strTempMDBFile, "idb_cond_input", str.Trim(), str2.Trim());
                //tree table
                str = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBTree, "Text", false);
                str2 = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBTreeTable, "Text", false);
                p_dao1.CreateTableLink(this.m_strTempMDBFile, "idb_tree_input", str.Trim(), str2.Trim());
                //site tree
                str = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBSiteTree, "Text", false);
                str2 = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBSiteTreeTable, "Text", false);
                p_dao1.CreateTableLink(this.m_strTempMDBFile, "idb_site_tree_input", str.Trim(), str2.Trim());

				//destroy the object and release it from memory
				p_dao1.m_DaoWorkspace.Close();
				p_dao1 = null;

                SetThermValue(m_frmTherm.progressBar1, "Value", 20);

				//get an ado connection string for the temp mdb file
				this.m_strTempMDBFileConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"","");

				//get the biosum plot, cond, and tree table names
				this.m_strPlotTable = p_datasource.getValidDataSourceTableName("PLOT");
				this.m_strCondTable = p_datasource.getValidDataSourceTableName("CONDITION");
				this.m_strTreeTable = p_datasource.getValidDataSourceTableName("TREE");
				this.m_strSiteTreeTable = p_datasource.getValidDataSourceTableName("SITE TREE");

   					
				//create a new connection to the temp MDB file
				this.m_connTempMDBFile = new System.Data.OleDb.OleDbConnection();

				//open the connection to the temp mdb file 
				this.m_ado.OpenConnection(this.m_strTempMDBFileConn,ref this.m_connTempMDBFile);

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				//get the fiabiosum table structures
				System.Data.DataTable p_dtPlotSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strPlotTable);
				System.Data.DataTable p_dtCondSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strCondTable);
				System.Data.DataTable p_dtTreeSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strTreeTable);
				System.Data.DataTable p_dtSiteTreeSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strSiteTreeTable);
				//get the idb table structures
				System.Data.DataTable p_dtIDBPlotSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from idb_plot_input");
				System.Data.DataTable p_dtIDBCondSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from idb_cond_input");
				System.Data.DataTable p_dtIDBTreeSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from idb_tree_input");
				System.Data.DataTable p_dtIDBSiteTreeSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from idb_site_tree_input");

				System.Data.DataTable p_dtTreeWorkTable = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from idb_tree_input");
				System.Data.DataTable p_dtSiteTreeWorkTable = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from idb_site_tree_input");



                SetThermValue(m_frmTherm.progressBar1, "Value", 30);
                

             
				this.m_strSQL = "SELECT biosum_plot_id, statecd as cond_ttl " + 
					"FROM " + this.m_strPlotTable.Trim() + ";";


				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dtPlotWorkTable = this.m_ado.getTableSchema(this.m_connTempMDBFile,this.m_strSQL);

				
                this.m_ado.m_strSQL = "SELECT * INTO plot_column_updates_work_table " + 
					                  "FROM (SELECT biosum_plot_id, statecd as cond_ttl " + 
											"FROM " + this.m_strPlotTable.Trim() + ") "  + 
					                  "WHERE 1 = 2";

				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);

				this.m_ado.m_strSQL = "SELECT biosum_plot_id INTO temp_plot_table FROM " + this.m_strPlotTable + " WHERE 1=2";
				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);

				this.m_ado.m_strSQL = "select * INTO temp_tree_input FROM idb_tree_input WHERE 1 = 2";
				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);

				this.m_ado.m_strSQL = "select * INTO tree_work_table FROM idb_tree_input WHERE 1 = 2";
				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);

                SetThermValue(m_frmTherm.progressBar1, "Value", 100);
                SetThermValue(m_frmTherm.progressBar2, "Value", 20);
                
				System.Threading.Thread.Sleep(2000);

				//-------------------------------PLOT----------------------------------//
                SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                SetLabelValue(m_frmTherm.lblMsg,"Text","Plot Table: Inserting Plot Records...Stand By");
                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

				//build field list string to insert sql by matching 
				//up the column names in the biosum plot table and the fiadb plot table
				strFields = "";
				for (x=0; x<=p_dtPlotSchema.Rows.Count-1;x++)
				{
					strCol = p_dtPlotSchema.Rows[x]["columnname"].ToString().Trim();
					//see if there is an equivalent FIADB column
					for (y=0; y<=p_dtIDBPlotSchema.Rows.Count-1;y++)
					{
						if (strCol.Trim().ToUpper() == p_dtIDBPlotSchema.Rows[y]["columnname"].ToString().ToUpper())
						{
							if (strFields.Trim().Length == 0)
							{
								strFields = strCol;
							}
							else
							{	
								strFields += "," + strCol;
							}
							break;
						}
					}
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 50);
                /********************************************************
				 **create plot input insert command
				 ********************************************************/
				//check the user defined filters
				if (Checked(rdoFilterNone)==true)
				{
					//forested/nonforested filters
					if (Checked(chkNonForested)==true && Checked(chkForested)==true)
					{
						//all plots
						this.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (" + strFields + ")" + 
							" SELECT " + strFields + " " + 
							" FROM idb_plot_input " + 
							" WHERE MID(biosum_plot_id,2,4)='" + this.m_strIDBInv.Trim() + "' ";
					}
					else if (Checked(chkForested)==true)
					{
						//forested plots
						this.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (" + strFields + ")" + 
							" SELECT " + strFields + " FROM idb_plot_input " + 
							" WHERE MID(biosum_plot_id,2,4)='" + this.m_strIDBInv.Trim() + "' AND " + 
							" Plot_status_cd = 1;";

					}
					else
					{
						//nonforested plots
						this.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (" + strFields + ")" + 
							" SELECT " + strFields + " FROM idb_plot_input " + 
							" WHERE MID(biosum_plot_id,2,4)='" + this.m_strIDBInv.Trim() + "' AND " + 
							"plot_status_cd IS NULL or plot_status_cd <> 1;";
					}
				}
				else if (Checked(rdoFilterByFile) == true)
				{
					//user defined plot_cn filter
					this.m_strSQL = "INSERT INTO  " + this.m_strPlotTable + " (" + strFields + ")" + 
						" SELECT " + strFields + " FROM idb_plot_input WHERE idb_plot_id  IN (" + this.m_strPlotIdList.Trim() + ");";
				}
				else if (Checked(rdoFilterByMenu)==true) 
				{
					if (Checked(chkNonForested)==true && Checked(chkForested)==true)
					{
						//all plots
						this.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (" + strFields + ")" + 
							" SELECT " + strFields + " FROM idb_plot_input " + 
							" WHERE MID(biosum_plot_id,2,4)='" + this.m_strIDBInv.Trim() + "' AND ";
					}
					else if (Checked(chkForested)==true)
					{
						//forested plots
						this.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (" + strFields + ")" + 
							" SELECT " + strFields + " FROM idb_plot_input " + 
							" WHERE MID(biosum_plot_id,2,4)='" + this.m_strIDBInv.Trim() + "' AND " + 
							       "(plot_status_cd = 1) AND ";
					}
					else
					{
						//nonforested plots
						this.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (" + strFields + ")" + 
							" SELECT " + strFields + " FROM fiadb_plot_input " + 
							" WHERE MID(biosum_plot_id,2,4)='" + this.m_strIDBInv.Trim() + "' AND " + 
							       "(plot_status_cd IS NULL or plot_status_cd <> 1) AND ";
					}
					if (this.m_strStateCountyPlotSQL.Trim().Length > 0)
					{
						//state,county,plot filter
						this.m_strSQL += this.m_strStateCountyPlotSQL.Trim() + ";";
					}
					else
					{
						//state,county filter
						this.m_strSQL += this.m_strStateCountySQL.Trim() + ";";
					}

				}
				else
				{
					this.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (" + strFields + ")" + 
						" SELECT " + strFields + " FROM fiadb_plot_input " + 
						" WHERE MID(biosum_plot_id,2,4)='" + this.m_strIDBInv.Trim() + "';";
				
				}
				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
                SetThermValue(m_frmTherm.progressBar1, "Value", 75);

				if (this.m_ado.m_intError == 0)
				{
					this.m_ado.m_strSQL="INSERT INTO temp_plot_table SELECT biosum_plot_id FROM " + this.m_strPlotTable + " WHERE biosum_status_cd=9";
					this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);

					intAddedPlotRows = Convert.ToInt32(this.m_ado.getRecordCount(this.m_connTempMDBFile,"select count(*) from " + this.m_strPlotTable + " where biosum_status_cd = 9;",this.m_strPlotTable));
				}
				else
				{
					this.m_intError=this.m_ado.m_intError;
					//error occured so remove new plot records
					this.m_strSQL = "DELETE FROM " + this.m_strPlotTable + " WHERE biosum_status_cd=9;";
					this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 100);
				System.Threading.Thread.Sleep(2000);
                SetThermValue(m_frmTherm.progressBar2, "Value", 40);
                SetThermValue(m_frmTherm.progressBar1, "Value", 0);
				//-------------------------------COND----------------------------------//
				if (intAddedPlotRows > 0 && this.m_intError==0 && this.m_ado.m_intError==0)
				{
                    SetLabelValue(m_frmTherm.lblMsg,"Text","Cond Table: Inserting Condition Records...Stand By");
					frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)m_frmTherm,"Refresh");

					//build field list string to insert sql by matching 
					//up the column names in the biosum cond table and the fiadb cond table
					strFields = "";
					strValues = "";
					for (x=0; x<=p_dtCondSchema.Rows.Count-1;x++)
					{
						strCol = p_dtCondSchema.Rows[x]["columnname"].ToString().Trim();
						//see if there is an equivalent FIADB column
						for (y=0; y<=p_dtIDBCondSchema.Rows.Count-1;y++)
						{
							if (strCol.Trim().ToUpper() == p_dtIDBCondSchema.Rows[y]["columnname"].ToString().ToUpper())
							{
								if (strFields.Trim().Length == 0)
								{
									strFields = strCol;
									strValues = "c." + strCol.Trim();
									//strValues = strCol;
								}
								else
								{	
									strFields += "," + strCol;
									strValues += ",c." + strCol.Trim();
									//strValues += "," + strCol.Trim();
								}
								break;
							}
						}
					}
                    SetThermValue(m_frmTherm.progressBar1, "Value", 40);
				
					//create cond input insert command
				
					

					this.m_strSQL = "INSERT INTO " + this.m_strCondTable + " (" + strFields + ")" + 
						" SELECT " + strValues + " FROM idb_cond_input c INNER JOIN " + this.m_strPlotTable + " p ON c.idb_plot_id=p.idb_plot_id WHERE p.biosum_status_cd=9";

					strTime = System.DateTime.Now.ToString();
				
					this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
                    SetThermValue(m_frmTherm.progressBar1, "Value", 80);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
					
					this.m_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " " + 
						     "SET owngrpcd = IIF(owngrpcd IS NOT NULL AND owngrpcd < 10, (owngrpcd * 10),owngrpcd)";

					this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);
                    SetThermValue(m_frmTherm.progressBar1, "Value", 85);

				

					if (this.m_ado.m_intError != 0)
					{
						//remove new plot and cond records since error occured
						this.m_intError=this.m_ado.m_intError;
						this.m_strSQL = "DELETE FROM " + this.m_strPlotTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strCondTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
					}

					//delete plot records from the input table
					//that already exist in the biosum plot table
					strTime = System.DateTime.Now.ToString();

				
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 100);
				System.Threading.Thread.Sleep(2000);
                SetThermValue(m_frmTherm.progressBar2, "Value", 60);
                SetThermValue(m_frmTherm.progressBar1, "Value", 0);
				//-------------------------------TREE----------------------------------//
				if (intAddedPlotRows > 0 && this.m_intError==0 && this.m_ado.m_intError==0)
				{
                    SetLabelValue(m_frmTherm.lblMsg,"Text","Tree Table: Inserting Tree Records...Stand By");
					frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
					
					strTime = System.DateTime.Now.ToString();
					//copy the inv trees to a temp file
					if (this.m_strPlotIdList.Trim().Length > 0)
					{
						this.m_strSQL = " INSERT INTO tree_work_table SELECT * FROM idb_tree_input";
					}
					else
					{
						this.m_strSQL = " INSERT INTO tree_work_table SELECT * FROM idb_tree_input WHERE MID(biosum_cond_id,2,4) = '" + this.m_strIDBInv.Trim() + "'";
					}
					this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);

                    SetThermValue(m_frmTherm.progressBar1, "Value", 10);

					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				
					/***************************************************
					 **get a list of fields that are not in the 
					 **biosum tree table and delete them from the
					 **tree work table
					 ***************************************************/
					string[] strFieldList;
					strFieldList = new string[p_dtIDBTreeSchema.Rows.Count];

					for (x=0; x<=p_dtIDBTreeSchema.Rows.Count-1;x++)
					{
						strFieldList[x]="";
						strCol = p_dtIDBTreeSchema.Rows[x]["columnname"].ToString().Trim();
						//see if there is an equivalent FIADB column
						for (y=0; y<=p_dtTreeSchema.Rows.Count-1;y++)
						{
							if (strCol.Trim().ToUpper() == p_dtTreeSchema.Rows[y]["columnname"].ToString().ToUpper())
							{
								break;
							}
						}
						if (y > p_dtTreeSchema.Rows.Count-1)
							strFieldList[x]=strCol;
					}
                    SetThermValue(m_frmTherm.progressBar1, "Value", 20);

					//close the connection to the temp mdb file
					this.m_connTempMDBFile.Close();
					while (m_connTempMDBFile.State == System.Data.ConnectionState.Open)
						System.Threading.Thread.Sleep(1000);
					m_connTempMDBFile.Dispose();
					m_connTempMDBFile=null;
					


					//instatiate dao again to delete the fields from the tree work table
					//in the temp MDB file
					dao_data_access p_dao3 = new dao_data_access();

					p_dao3.DeleteField(this.m_strTempMDBFile,"tree_work_table", strFieldList);

					p_dao3 = null;
                    SetThermValue(m_frmTherm.progressBar1, "Value",30);
					

					

					//reopen the connection to the temp mdb file 
					this.m_connTempMDBFile = new System.Data.OleDb.OleDbConnection();
					this.m_ado.OpenConnection(this.m_strTempMDBFileConn,ref this.m_connTempMDBFile);

					strTime = System.DateTime.Now.ToString();
					//copy the inv trees to a temp file
					this.m_strSQL = " INSERT INTO " + this.m_strTreeTable + " SELECT t.* FROM tree_work_table t INNER JOIN " + this.m_strCondTable + " c ON t.biosum_cond_id = c.biosum_cond_id WHERE c.biosum_status_cd=9;";
					this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
                    SetThermValue(m_frmTherm.progressBar1, "Value", 75);
					if (this.m_ado.m_intError != 0)
					{
						//remove new plot and cond records since error occured
						this.m_intError=this.m_ado.m_intError;
						this.m_strSQL = "DELETE FROM " + this.m_strPlotTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strCondTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strTreeTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
					}
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 100);
				System.Threading.Thread.Sleep(2000);
                SetThermValue(m_frmTherm.progressBar2, "Value", 80);
                SetThermValue(m_frmTherm.progressBar1, "Value", 0);
				//-------------------------------SITE TREE----------------------------------//
				if (intAddedPlotRows > 0 && this.m_intError==0 && this.m_ado.m_intError==0)
				{
                    SetLabelValue(m_frmTherm.lblMsg,"Text","Site Tree Table: Inserting Site Tree Records...Stand By");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");


					//build field list string to insert sql by matching 
					//up the column names in the biosum cond table and the fiadb cond table
					
					strFields = "";
					strValues = "";
					for (x=0; x<=p_dtSiteTreeSchema.Rows.Count-1;x++)
					{
						strCol = p_dtSiteTreeSchema.Rows[x]["columnname"].ToString().Trim();
						//see if there is an equivalent FIADB column
						for (y=0; y<=p_dtIDBSiteTreeSchema.Rows.Count-1;y++)
						{
							if (strCol.Trim().ToUpper() == p_dtIDBSiteTreeSchema.Rows[y]["columnname"].ToString().ToUpper())
							{
								if (strFields.Trim().Length == 0)
								{
									strFields = strCol;
									strValues = "t." + strCol.Trim();
								}
								else
								{	
									strFields += "," + strCol;
									strValues += ",t." + strCol.Trim();
								}
								break;
							}
						}
					}
                    SetThermValue(m_frmTherm.progressBar1, "Value", 50);
					//create cond input insert command

					this.m_strSQL = "INSERT INTO " + this.m_strSiteTreeTable + " (" + strFields + ")" + 
						" SELECT " + strValues + " FROM idb_site_tree_input t INNER JOIN temp_plot_table p ON TRIM(t.biosum_plot_id)=p.biosum_plot_id";

					strTime = System.DateTime.Now.ToString();
				
					this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);

					if (this.m_ado.m_intError != 0)
					{
						//remove new plot and cond records since error occured
						this.m_intError=this.m_ado.m_intError;
						this.m_strSQL = "DELETE FROM " + this.m_strPlotTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strCondTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strTreeTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strSiteTreeTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
					}


					

					
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 100);
				System.Threading.Thread.Sleep(2000);
                SetThermValue(m_frmTherm.progressBar2, "Value", 95);
                SetThermValue(m_frmTherm.progressBar1, "Value", 0);
				if (intAddedPlotRows > 0 && this.m_intError==0 && this.m_ado.m_intError==0)
				{
				    SetLabelValue(m_frmTherm.lblMsg,"Text","Updating Columns...Stand By");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
				

					this.UpdateColumns(this.m_ado);
					

					if (this.m_intError==0)
					{
				

						intAddedPlotRows= (int)this.m_ado.getRecordCount(this.m_connTempMDBFile,"select count(*) from " + this.m_strPlotTable + " WHERE biosum_status_cd=9",this.m_strPlotTable);
						intAddedCondRows= (int)this.m_ado.getRecordCount(this.m_connTempMDBFile,"select count(*) from " + this.m_strCondTable  + " WHERE biosum_status_cd=9",this.m_strCondTable);
						intAddedTreeRows= (int)this.m_ado.getRecordCount(this.m_connTempMDBFile,"select count(*) from " + this.m_strTreeTable  + " WHERE biosum_status_cd=9",this.m_strTreeTable);
						intAddedSiteTreeRows= (int)this.m_ado.getRecordCount(this.m_connTempMDBFile,"select count(*) from " + this.m_strSiteTreeTable  + " WHERE biosum_status_cd=9",this.m_strSiteTreeTable);


						this.m_strSQL = " UPDATE " + this.m_strPlotTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = " UPDATE " + this.m_strCondTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = " UPDATE " + this.m_strTreeTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = " UPDATE " + this.m_strSiteTreeTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);


                        SetThermValue(m_frmTherm.progressBar1, "Value",GetThermValue(m_frmTherm.progressBar1,"Maximum"));
                        SetLabelValue(m_frmTherm.lblMsg,"Text","Done");
                        frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

						MessageBox.Show("Successfully Appended \n" + 
							intAddedPlotRows.ToString().Trim() + " Plot Records \n" + 
							intAddedCondRows.ToString().Trim() + " Condition Records \n" + 
							intAddedTreeRows.ToString().Trim() + " Tree Records \n" + 
							intAddedSiteTreeRows.ToString().Trim() + " Site Tree Records","Add Plot Data");

					}
					else
					{

						//error occurred in the updatecolumns so delete the records
						this.m_strSQL = "DELETE FROM " + this.m_strPlotTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						//delete added cond records since error occured
						this.m_strSQL = "DELETE FROM " + this.m_strCondTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						//delete added cond records since error occured
						this.m_strSQL = "DELETE FROM " + this.m_strTreeTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strSiteTreeTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);

                        SetThermValue(m_frmTherm.progressBar1, "Value", GetThermValue(m_frmTherm.progressBar1, "Maximum"));
                        frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

					}
				}
				this.m_strCurrentProcess="";
                SetThermValue(m_frmTherm.progressBar1, "Value", GetThermValue(m_frmTherm.progressBar1, "Maximum"));
                SetThermValue(m_frmTherm.progressBar2, "Value", 100);		
				if (this.m_intError != 0 || this.m_ado.m_intError != 0)
					MessageBox.Show("!!Error Occured Adding Plot Records: 0 Records Added!!","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				
				if (this.m_connTempMDBFile.State == System.Data.ConnectionState.Open)
				{
					this.m_connTempMDBFile.Close();
					while (m_connTempMDBFile.State == System.Data.ConnectionState.Open)
						System.Threading.Thread.Sleep(1000);
					m_connTempMDBFile.Dispose();
					m_connTempMDBFile=null;
				}
				this.m_ado=null;
                
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Visible", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);


                LoadMDBPlotCondTreeData_Finish();

            }
            catch (System.Threading.ThreadInterruptedException err)
            {
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (System.Threading.ThreadAbortException err)
            {
                if (this.m_connTempMDBFile != null)
                {
                    if (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
                    {
                        m_ado.CloseConnection(m_connTempMDBFile);
                    }
                    m_connTempMDBFile = null;
                }
                if (m_ado != null)
                {
                    if (m_ado.m_DataSet != null)
                    {
                        this.m_ado.m_DataSet.Clear();
                        this.m_ado.m_DataSet.Dispose();
                    }
                    this.m_ado = null;
                }
                this.CancelThreadCleanup();
                this.ThreadCleanUp();
                this.CleanupThread();

            }
            catch (Exception err)
            {
                MessageBox.Show("!!Error!! \n" +
                    "Module - uc_plot_input:LoadIDBPlotCondTreeData_Process  \n" +
                    "Err Msg - " + err.Message.ToString().Trim(),
                    "FVS Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }
            finally
            {

            }

            if (this.m_connTempMDBFile != null)
            {
                if (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
                {
                    m_ado.CloseConnection(m_connTempMDBFile);
                }
                m_connTempMDBFile = null;
            }
            if (m_ado != null)
            {
                if (m_ado.m_DataSet != null)
                {
                    this.m_ado.m_DataSet.Clear();
                    this.m_ado.m_DataSet.Dispose();
                }
                this.m_ado = null;
            }
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);
            if (this.m_frmTherm != null) frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", false);




            LoadMDBPlotCondTreeData_Finish();

            CleanupThread();

            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
			
		
		
		}
        private bool Checked(System.Windows.Forms.RadioButton p_rdoButton)
        {
            return (bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.RadioButton)p_rdoButton, "Checked", false);
        }
        private bool Checked(System.Windows.Forms.CheckBox p_chkBox)
        {
            return (bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.CheckBox)p_chkBox, "Checked", false);
        }
        private void SetThermValue(System.Windows.Forms.ProgressBar p_oPb,string p_strPropertyName, int p_intValue)
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)p_oPb,p_strPropertyName,(int)p_intValue);
        }
        private int GetThermValue(System.Windows.Forms.ProgressBar p_oPb, string p_strPropertyName)
        {
            return (int)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)p_oPb, p_strPropertyName, false);
        }
        private bool GetBooleanValue(System.Windows.Forms.Control p_oControl, string p_strPropertyName)
        {
            return (bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)p_oControl,p_strPropertyName,false);
        }
            
            
            
        private void SetLabelValue(System.Windows.Forms.Label p_oLabel, string p_strPropertyName, string p_strValue)
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Label)p_oLabel, p_strPropertyName, p_strValue);
        }

		private void StartTherm(string p_strNumberOfTherms,string p_strTitle)
		{
			this.m_frmTherm = new frmTherm((frmDialog)this.ParentForm, p_strTitle);

			this.m_frmTherm.Text = p_strTitle;
			this.m_frmTherm.lblMsg.Text="";
			this.m_frmTherm.lblMsg2.Text="";
			this.m_frmTherm.Visible=false;
			this.m_frmTherm.btnCancel.Visible=true;
			this.m_frmTherm.lblMsg.Visible=true;
			this.m_frmTherm.progressBar1.Minimum=0;
			this.m_frmTherm.progressBar1.Visible=true;
			this.m_frmTherm.progressBar1.Maximum = 10;

			if (p_strNumberOfTherms=="2")
			{
				this.m_frmTherm.progressBar2.Size = this.m_frmTherm.progressBar1.Size;
				this.m_frmTherm.progressBar2.Left = this.m_frmTherm.progressBar1.Left;
				this.m_frmTherm.progressBar2.Top = Convert.ToInt32(this.m_frmTherm.progressBar1.Top + (this.m_frmTherm.progressBar1.Height * 3));
				this.m_frmTherm.lblMsg2.Top = this.m_frmTherm.progressBar2.Top + this.m_frmTherm.progressBar2.Height + 5;
				this.m_frmTherm.Height = this.m_frmTherm.lblMsg2.Top + this.m_frmTherm.lblMsg2.Height + this.m_frmTherm.btnCancel.Height + 50;
				this.m_frmTherm.btnCancel.Top = this.m_frmTherm.ClientSize.Height - this.m_frmTherm.btnCancel.Height - 5;
				this.m_frmTherm.lblMsg2.Show();
				this.m_frmTherm.progressBar2.Visible=true;
			}
			this.m_frmTherm.AbortProcess = false;
			this.m_frmTherm.Refresh();
			this.m_frmTherm.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			//((frmDialog)this.ParentForm).Enabled=false;
			this.m_frmTherm.Visible=true;
			
		}
        public void StopThread()
        {

            string strMsg = "";
            
            frmMain.g_oDelegate.AbortProcessing("FIA Biosum", "Do you wish to cancel adding plot data?");

            if (frmMain.g_oDelegate.CurrentThreadProcessAborted)
            {
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm,"AbortProcess",true);
                this.CancelThreadCleanup();
                this.ThreadCleanUp();
            }


        }
		
		public void AddPlotRecordsFinished()
		{
			this.m_strPlotIdList="";
			this.thdProcessRecords.Abort();
			if (this.thdProcessRecords.IsAlive)
			{
				frmMain.g_oDelegate.SetControlPropertyValue(
                    (System.Windows.Forms.Label)m_frmTherm.lblMsg,"Text","Attempting To Abort Process...Stand By");
                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Label)this.m_frmTherm.lblMsg, "Refresh");
				this.thdProcessRecords.Join(1000);
			}
			this.thdProcessRecords = null;
		}
		private void CancelThreadCleanup()
		{
			try
			{

				if (this.m_ado == null)
				{
					this.m_ado = new ado_data_access();
				}
				else
				{

					this.m_ado = null;
					this.m_ado = new ado_data_access();
				}
				if (this.m_connTempMDBFile == null)
				{

				}
				else
				{

					this.m_connTempMDBFile = null;
				}
				this.m_ado.OpenConnection(this.m_strTempMDBFileConn);


				if (this.m_strCurrentProcess=="mdbFIADBFileInput" ||
					this.m_strCurrentProcess=="txtFileInput" ||
					this.m_strCurrentProcess=="mdbIDBFileInput")
				{
					int intAddedPlotRows= (int)this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection,"select count(*) from " + this.m_strPlotTable + " WHERE biosum_status_cd=9",this.m_strPlotTable);
					if (intAddedPlotRows > 0)
					{
						//error occurred in the updatecolumns so delete the records
						this.m_strSQL = "DELETE FROM " + this.m_strPlotTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_strSQL);

					}
					int intAddedCondRows= (int)this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection,"select count(*) from " + this.m_strCondTable + " WHERE biosum_status_cd=9",this.m_strCondTable);
					if (intAddedCondRows > 0)
					{
						this.m_strSQL = "DELETE FROM " + this.m_strCondTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_strSQL);

					}

					int intAddedTreeRows= (int)this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection,"select count(*) from " + this.m_strTreeTable + " WHERE biosum_status_cd=9",this.m_strTreeTable);
					if (intAddedTreeRows > 0)
					{
					
						//delete added tree records since error occured
						this.m_strSQL = "DELETE FROM " + this.m_strTreeTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_strSQL);
					

					}
					int intAddedSiteTreeRows= (int)this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection,"select count(*) from " + this.m_strSiteTreeTable + " WHERE biosum_status_cd=9",this.m_strSiteTreeTable);
					if (intAddedSiteTreeRows > 0)
					{
					
						//delete added tree records since error occured
						this.m_strSQL = "DELETE FROM " + this.m_strSiteTreeTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_strSQL);
					}
				}
				if ((bool)frmMain.g_oDelegate.GetControlPropertyValue(
                      (System.Windows.Forms.RadioButton)rdoFIADB,"Checked",false) && 
                      (this.m_strCurrentProcess=="txtFileInputPopFiles" || 
					   this.m_strCurrentProcess=="mdbFiadbInputPopTables"))
				{
					
					int intAddedPopEvalRows= (int)this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection,"select count(*) from " + this.m_strPopEvalTable + " WHERE biosum_status_cd=9",this.m_strPopEvalTable);
					if (intAddedPopEvalRows > 0)
					{
					
						//delete added tree records since error occured
						this.m_strSQL = "DELETE FROM " + this.m_strPopEvalTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_strSQL);
						this.m_strLoadedPopEvalTxtInputFile="";
					

					}
					int intAddedPopStratumRows= (int)this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection,"select count(*) from " + this.m_strPopStratumTable + " WHERE biosum_status_cd=9",this.m_strPopStratumTable);
					if (intAddedPopStratumRows > 0)
					{
					
						//delete added tree records since error occured
						this.m_strSQL = "DELETE FROM " + m_strPopStratumTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_strSQL);
						this.m_strLoadedPopStratumTxtInputFile="";
					

					}

					int intAddedPopEstUnitRows= (int)this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection,"select count(*) from " + this.m_strPopEstUnitTable + " WHERE biosum_status_cd=9",this.m_strPopEstUnitTable);
					if (intAddedPopEstUnitRows > 0)
					{
					
						//delete added tree records since error occured
						this.m_strSQL = "DELETE FROM " + m_strPopEstUnitTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_strSQL);
						this.m_strLoadedPopEstUnitTxtInputFile="";
					

					}
					int intAddedPpsaRows= (int)this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection,"select count(*) from " + this.m_strPpsaTable + " WHERE biosum_status_cd=9",this.m_strPpsaTable);
					if (intAddedPpsaRows > 0)
					{
					
						//delete added tree records since error occured
						this.m_strSQL = "DELETE FROM " + this.m_strPpsaTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_strSQL);
						this.m_strLoadedPpsaTxtInputFile="";
					}
					
				}
				MessageBox.Show("!!User Canceled Adding Plot Records: 0 Records Added!!","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm,"Visible",false);
				//((frmDialog)this.ParentForm).m_frmMain.Visible=true;
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Visible", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);



			}
			catch 
			{
			}

		}

		private void btnTreeRegionalBiomassBrowse_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "FIADB Tree Regional Biomass Table Data";
			OpenFileDialog1.Filter = "Comma Delimited Text File (*.CSV;*.TXT;*.DAT) |*.csv;*.txt;*.dat";
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK) 
			{
				if (OpenFileDialog1.FileName.Trim().Length > 0) 
				{
					m_strTreeRegionalBiomassTxtInputFile = OpenFileDialog1.FileName.Trim();
					this.txtTreeRegionalBiomass.Text = m_strTreeRegionalBiomassTxtInputFile;
				}
			}
			else 
			{
			}
			OpenFileDialog1 = null;
		}

		private void btnPopEvalBrowse_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "FIADB Population Evaluation Table Data";
			OpenFileDialog1.Filter = "Comma Delimited Text File (*.CSV;*.TXT;*.DAT) |*.csv;*.txt;*.dat";
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK) 
			{
				if (OpenFileDialog1.FileName.Trim().Length > 0) 
				{
					this.m_strPopEvalTxtInputFile = OpenFileDialog1.FileName.Trim();
					this.txtPopEval.Text = this.m_strPopEvalTxtInputFile;
				}
			}
			else 
			{
			}
			OpenFileDialog1 = null;
		}

		private void btnPopEstUnitBrowse_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "FIADB Population Estimation Unit Table Data";
			OpenFileDialog1.Filter = "Comma Delimited Text File (*.CSV;*.TXT;*.DAT) |*.csv;*.txt;*.dat";
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK) 
			{
				if (OpenFileDialog1.FileName.Trim().Length > 0) 
				{
					this.m_strPopEstUnitTxtInputFile = OpenFileDialog1.FileName.Trim();
					this.txtPopEstUnit.Text = this.m_strPopEstUnitTxtInputFile;
				}
			}
			else 
			{
			}
			OpenFileDialog1 = null;
		}

		private void btnPopStratumBrowse_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "FIADB Population Stratum Table Data";
			OpenFileDialog1.Filter = "Comma Delimited Text File (*.CSV;*.TXT;*.DAT) |*.csv;*.txt;*.dat";
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK) 
			{
				if (OpenFileDialog1.FileName.Trim().Length > 0) 
				{
					this.m_strPopStratumTxtInputFile = OpenFileDialog1.FileName.Trim();
					this.txtPopStratum.Text = this.m_strPopStratumTxtInputFile;
				}
			}
			else 
			{
			}
			OpenFileDialog1 = null;
		}

		private void btnPpsaBrowse_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "FIADB Population Plot Stratum Assignment Table Data";
			OpenFileDialog1.Filter = "Comma Delimited Text File (*.CSV;*.TXT;*.DAT) |*.csv;*.txt;*.dat";
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK) 
			{
				if (OpenFileDialog1.FileName.Trim().Length > 0) 
				{
					this.m_strPpsaTxtInputFile = OpenFileDialog1.FileName.Trim();
					this.txtPpsa.Text = this.m_strPpsaTxtInputFile;
				}
			}
			else 
			{
			}
			OpenFileDialog1 = null;
		}

		private void btnFIADBInvAppend_Click(object sender, System.EventArgs e)
		{
            m_intError = 0;
			((frmDialog)this.ParentForm).MinimizeMainForm=true;
			this.Enabled=false;
			if (this.lstFIADBInv.SelectedItems.Count > 0)
			{
				if (this.rdoAccess.Checked)
				{
                    this.CalculateAdjustments_Start();
                    if (m_intError == 0)
                    {
                        if (m_intError == 0)
                        {
                            this.LoadMDBFiadbPopFiles();
                        }
                        this.m_strLoadedPopEstUnitTxtInputFile = "";
                        this.m_strLoadedPopEvalTxtInputFile = "";
                        this.m_strLoadedPopStratumTxtInputFile = "";
                        this.m_strLoadedPpsaTxtInputFile = "";
                        if (this.rdoFilterNone.Checked && m_intError == 0)
                        {

                            this.LoadMDBPlotCondTreeData_Start();

                        }
                        else if (this.rdoFilterByFile.Checked && m_intError == 0)
                        {
                            this.m_strPlotIdList = this.CreateDelimitedStringList(this.txtFilterByFile.Text.Trim(), ",", ",", false);
                            if (this.m_intError == 0)
                                this.LoadMDBPlotCondTreeData_Start();

                        }
                    }
					

				}
				else
				{
					LoadTxtPopFiles();
					if (this.rdoFilterNone.Checked && m_intError==0)
					{

                        LoadTxtPlotCondTreeData_Start();
					}
					else if (this.rdoFilterByFile.Checked && m_intError==0)
					{
						this.m_strPlotIdList = this.CreateDelimitedStringList(this.txtFilterByFile.Text.Trim(), ",", "," , false);
						if (this.m_intError==0)
                            LoadTxtPlotCondTreeData_Start();

					}
				}
				

			}
			else
			{
				MessageBox.Show("Select an FIADB population evaluation","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			}
			//((frmDialog)this.ParentForm).MinimizeMainForm=false;
			//this.Enabled=true;
		}

		private void btnFIADBInvPrevious_Click(object sender, System.EventArgs e)
		{
			this.grpboxFIADBInv.Hide();
			this.grpboxFilter.Show();
		}

		private void btnFIADBInvCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}

		private void btnFIADBInvNext_Click(object sender, System.EventArgs e)
		{
			if (this.lstFIADBInv.SelectedItems.Count > 0)
			{
				if (this.rdoAccess.Checked)
				{
					if (this.m_strLoadedPopEstUnitInputTable.Trim().ToUpper() !=
						this.cmbFiadbPopEstUnitTable.Text.Trim().ToUpper() ||
						this.m_strLoadedPopStratumInputTable.Trim().ToUpper() !=
						this.cmbFiadbPopStratumTable.Text.Trim().ToUpper() ||
						this.m_strLoadedPpsaInputTable.Trim().ToUpper() !=
						this.cmbFiadbPpsaTable.Text.Trim().ToUpper() ||
						this.m_strLoadedFIADBEvalId.Trim() != 
						this.m_strCurrFIADBEvalId.Trim() ||
						this.m_strLoadedFIADBRsCd.Trim() !=
						this.m_strCurrFIADBRsCd.Trim() ||
						this.m_strLoadedFiadbInputFile.Trim().ToUpper() != 
						this.txtMDBFiadbInputFile.Text.Trim().ToUpper())
					{
						this.LoadMDBFiadbPopFiles();
					}
				

					if (m_intError==0)
					{
						this.mdbInputStateCounty();
						this.grpboxFilterByState.Visible=true;
						this.lstFilterByState.Refresh();
						this.grpboxFIADBInv.Visible=false;

						//if (this.rdoFIADB.Checked==true && 
						//	this.rdoText.Checked==true && 
						//	this.rdoText.Enabled==true &&
						//	this.rdoFilterByMenu.Checked==true)
						//{
						//	this.txtFileInputStateCounty();
						//	this.grpboxFilterByState.Visible=true;
						//	this.lstFilterByState.Refresh();
						//	this.grpboxFIADBInv.Visible=false;

						//}
					}
				}
				else
				{
					//check if there have been any changes
					if (this.m_strLoadedPopEstUnitTxtInputFile.Trim().ToUpper() !=
						this.txtPopEstUnit.Text.Trim().ToUpper() ||
						this.m_strLoadedPopStratumTxtInputFile.Trim().ToUpper() !=
						this.txtPopStratum.Text.Trim().ToUpper() ||
						this.m_strLoadedPpsaTxtInputFile.Trim().ToUpper() !=
						this.txtPpsa.Text.Trim().ToUpper() ||
						this.m_strLoadedFIADBEvalId.Trim() != 
						this.m_strCurrFIADBEvalId.Trim() ||
						this.m_strLoadedFIADBRsCd.Trim() !=
						this.m_strCurrFIADBRsCd.Trim())
					{
						LoadTxtPopFiles();

					}
				

					if (m_intError==0)
					{
						if (this.rdoFIADB.Checked==true && 
							this.rdoText.Checked==true && 
							this.rdoText.Enabled==true &&
							this.rdoFilterByMenu.Checked==true)
						{
							this.txtFileInputStateCounty();
							this.grpboxFilterByState.Visible=true;
							this.lstFilterByState.Refresh();
							this.grpboxFIADBInv.Visible=false;

						}
					}
				}

			}
			else
			{
				MessageBox.Show("Select an FIADB population evaluation","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			}

		}
		private void LoadTxtPopFiles()
		{
			this.m_strCurrentProcess="txtFileInputPopFiles";
			this.m_bLoadStateCountyList=true;
			this.m_bLoadStateCountyPlotList=true;

			this.StartTherm("2","Add Text File Pop Table Data");
			this.m_frmTherm.progressBar2.Maximum=3;
			this.m_frmTherm.progressBar2.Minimum=0;
			this.m_frmTherm.progressBar2.Value=0;
			this.m_frmTherm.lblMsg2.Text = "Overall Progress";
			this.m_strTableType="POPULATION STRATUM";
			this.m_strCurrentTxtInputFile = this.txtPopStratum.Text;
			this.thdProcessRecords = new Thread(new ThreadStart(this.txtFileInputPopFiles));
			this.thdProcessRecords.IsBackground = true;
			this.thdProcessRecords.Start();
			while (thdProcessRecords.IsAlive)
			{
				thdProcessRecords.Join(1000);
				System.Windows.Forms.Application.DoEvents();

			}
			this.m_frmTherm.progressBar2.Value=1;
			thdProcessRecords=null;
			if (m_intError==0)
			{
				this.m_strLoadedPopStratumTxtInputFile = this.m_strCurrentTxtInputFile;
				this.m_strTableType="POPULATION ESTIMATION UNIT";
				this.m_frmTherm.lblMsg.Text = "pop estimation unit table";
				this.m_strCurrentTxtInputFile = this.txtPopEstUnit.Text;
				this.thdProcessRecords = new Thread(new ThreadStart(this.txtFileInputPopFiles));
				this.thdProcessRecords.IsBackground = true;
				this.thdProcessRecords.Start();
				while (thdProcessRecords.IsAlive)
				{
					thdProcessRecords.Join(1000);
					System.Windows.Forms.Application.DoEvents();

				}
				thdProcessRecords=null;

			}
			this.m_frmTherm.progressBar2.Value=2;
			if (m_intError==0)
			{
				this.m_strLoadedPopEstUnitTxtInputFile = this.m_strCurrentTxtInputFile;
				this.m_strTableType="POPULATION PLOT STRATUM ASSIGNMENT";
				this.m_frmTherm.lblMsg.Text = "ppsa table";
				this.m_strCurrentTxtInputFile = this.txtPpsa.Text;
				this.thdProcessRecords = new Thread(new ThreadStart(this.txtFileInputPopFiles));
				this.thdProcessRecords.IsBackground = true;
				this.thdProcessRecords.Start();
				while (thdProcessRecords.IsAlive)
				{
					thdProcessRecords.Join(1000);
					System.Windows.Forms.Application.DoEvents();

				}
				thdProcessRecords=null;

			}
			if (this.m_intError==0)
			{
				this.m_strLoadedPpsaTxtInputFile=this.m_strCurrentTxtInputFile;
			}
			this.m_frmTherm.progressBar2.Value=this.m_frmTherm.progressBar2.Maximum;
			System.Threading.Thread.Sleep(2000);
			this.m_frmTherm.Close();
			this.m_frmTherm = null;
			this.m_strCurrentProcess="";

		}
		private void LoadTxtPlotCondTreeData()
		{

		}
        private void CalculateAdjustments_Start()
        {
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oEventStopThread.Reset();
            frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            this.m_strCurrentProcess = "mdbFIADBFileInput";
            this.StartTherm("2", "Calculate Adjustment Factors");
            frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(CalculateAdjustments_Process));
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.m_oThread.Start();
            while (frmMain.g_oDelegate.m_oThread != null && frmMain.g_oDelegate.m_oThread.IsAlive)
            {
                frmMain.g_oDelegate.m_oThread.Join(1000);
                System.Windows.Forms.Application.DoEvents();

            }
        }
        private void LoadTxtPlotCondTreeData_Start()
        {
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oEventStopThread.Reset();
            frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            this.m_strCurrentProcess = "txtFileInput";
            this.StartTherm("2", "Add Text File Plot,Cond,Site Tree, & Tree Table Data");
            frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(LoadTxtPlotCondTreeData_Process));
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.m_oThread.Start();
        }

        private void LoadMDBPlotCondTreeData_Start()
        {
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oEventStopThread.Reset();
            frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            this.m_strCurrentProcess = "mdbFIADBFileInput";
            this.StartTherm("2", "Add MS Access Plot,Cond,Site Tree, & Tree Table Data");
            frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(LoadMDBPlotCondTreeData_Process));
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.m_oThread.Start();
        }
        private void LoadIDBPlotCondTreeData_Start()
        {
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oEventStopThread.Reset();
            frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            this.m_strCurrentProcess = "mdbIDBFileInput";
            this.StartTherm("2", "Add MS Access Plot,Cond,Site Tree, & Tree Table Data");
            frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(LoadIDBPlotCondTreeData_Process));
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.m_oThread.Start();
        }

        private void CalculateAdjustments_Finish()
        {
           

            if (this.m_frmTherm != null)
            {
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Close");
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Dispose");
                this.m_frmTherm = null;
            }
           
            this.m_strCurrentProcess = "";
        }

        private void LoadMDBPlotCondTreeData_Finish()
        {
            this.m_strPlotIdList = "";
            
            
            if (this.m_frmTherm != null)
            {
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Close");
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Dispose");
                this.m_frmTherm = null;
            }
            if (m_intError != 0)
            {
                this.m_strLoadedPopEstUnitInputTable = "";
                this.m_strLoadedPopStratumInputTable = "";
                this.m_strLoadedPpsaInputTable = "";
                this.m_strLoadedFiadbInputFile = "";
            }
            else
            {
                this.m_strLoadedPopEstUnitTxtInputFile = "";
                this.m_strLoadedPopEvalTxtInputFile = "";
                this.m_strLoadedPopStratumTxtInputFile = "";
                this.m_strLoadedPpsaTxtInputFile = "";
            }
            this.m_strCurrentProcess = "";
            frmMain.g_oDelegate.SetControlPropertyValue(this, "Enabled", true);
            ((frmDialog)this.ParentForm).MinimizeMainForm = false;
            
        }
		private void lstFilterByState_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
		{
			this.m_bLoadStateCountyPlotList=true;
		}

		private void btnMDBFiadbInputPrev_Click(object sender, System.EventArgs e)
		{
			this.grpboxMDBFiadbInput.Visible=false;
			this.grpboxInvType.Visible=true;
		}

		private void btnboxMDBFiadbInputFile_Click(object sender, System.EventArgs e)
		{
			
				OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
				OpenFileDialog1.Title = "MS Access Database File Containing FIADB Tables";
				OpenFileDialog1.Filter = "Microsoft Access Database File (*.MDB,*.MDE,*.ACCDB) |*.mdb;*.mde;*.accdb";
			
				DialogResult result =  OpenFileDialog1.ShowDialog();
				if (result == DialogResult.OK) 
				{
					if (OpenFileDialog1.FileName.Trim().Length > 0) 
					{
						string strFullPath = OpenFileDialog1.FileName.Trim();
						if (strFullPath.Length > 0) 
						{
							this.txtMDBFiadbInputFile.Text = strFullPath;
							dao_data_access tempDao = new dao_data_access();
							tempDao.OpenDb(strFullPath);
							if (tempDao.m_intError == 0) 
							{
								this.cmbFiadbCondTable.Items.Clear();
								this.cmbFiadbPlotTable.Items.Clear();
								this.cmbFiadbPopEstUnitTable.Items.Clear();
								this.cmbFiadbPopEvalTable.Items.Clear();
								this.cmbFiadbPopStratumTable.Items.Clear();
								this.cmbFiadbPpsaTable.Items.Clear();
								this.cmbFiadbTreeRegionalBiomassTable.Items.Clear();
								this.cmbFiadbTreeTable.Items.Clear();
								this.cmbFiadbSiteTreeTable.Items.Clear();
								for (int x=0; x <= tempDao.m_DaoDatabase.TableDefs.Count - 1; x++)
								{
									
									
									if (tempDao.m_DaoDatabase.TableDefs[x].Name.IndexOf("MSys",0) < 0) 		
									{
										this.cmbFiadbCondTable.Items.Add(tempDao.m_DaoDatabase.TableDefs[x].Name);
										this.cmbFiadbPlotTable.Items.Add(tempDao.m_DaoDatabase.TableDefs[x].Name);
										this.cmbFiadbPopEstUnitTable.Items.Add(tempDao.m_DaoDatabase.TableDefs[x].Name);
										this.cmbFiadbPopEvalTable.Items.Add(tempDao.m_DaoDatabase.TableDefs[x].Name);
										this.cmbFiadbPopStratumTable.Items.Add(tempDao.m_DaoDatabase.TableDefs[x].Name);
										this.cmbFiadbPpsaTable.Items.Add(tempDao.m_DaoDatabase.TableDefs[x].Name);
										this.cmbFiadbTreeRegionalBiomassTable.Items.Add(tempDao.m_DaoDatabase.TableDefs[x].Name);
										this.cmbFiadbTreeTable.Items.Add(tempDao.m_DaoDatabase.TableDefs[x].Name);
										this.cmbFiadbSiteTreeTable.Items.Add(tempDao.m_DaoDatabase.TableDefs[x].Name);
										switch (tempDao.m_DaoDatabase.TableDefs[x].Name.ToString().Trim().ToUpper())
										{
											case "COND":
												this.cmbFiadbCondTable.Text = "COND";
												break;
											case "TREE":
												this.cmbFiadbTreeTable.Text = "TREE";
												break;
											case "PLOT":
												this.cmbFiadbPlotTable.Text = "PLOT";
												break;
											case "POP_EVAL":
												this.cmbFiadbPopEvalTable.Text = "POP_EVAL";
												break;
											case "POP_ESTN_UNIT":
												this.cmbFiadbPopEstUnitTable.Text = "POP_ESTN_UNIT";
												break;
											case "POP_PLOT_STRATUM_ASSGN":
												this.cmbFiadbPpsaTable.Text = "POP_PLOT_STRATUM_ASSGN";
												break;
											case "POP_STRATUM":
												this.cmbFiadbPopStratumTable.Text = "POP_STRATUM";
												break;
											case "TREE_REGIONAL_BIOMASS":
												this.cmbFiadbTreeRegionalBiomassTable.Text = "TREE_REGIONAL_BIOMASS";
												break;
											case "SITETREE":
												this.cmbFiadbSiteTreeTable.Text = "SITETREE";
												break;

										}
									}

								}
                        
						
								tempDao.m_DaoDatabase.Close();
								tempDao.m_DaoDatabase = null;
							}
							tempDao = null;
						}
					}
				}
				else 
				{
				}
				OpenFileDialog1 = null;

			
		}

		private void btnMDBFiadbInputNext_Click(object sender, System.EventArgs e)
		{
			if (this.cmbFiadbPopEvalTable.Text.Trim().Length == 0 ||
				this.cmbFiadbCondTable.Text.Trim().Length == 0 ||
				this.cmbFiadbPlotTable.Text.Trim().Length == 0 ||
				this.cmbFiadbPopEstUnitTable.Text.Trim().Length == 0 ||
				this.cmbFiadbPopEvalTable.Text.Trim().Length == 0 ||
				this.cmbFiadbPopStratumTable.Text.Trim().Length == 0 ||
				this.cmbFiadbPpsaTable.Text.Trim().Length == 0 ||
				this.cmbFiadbTreeRegionalBiomassTable.Text.Trim().Length == 0 ||
				this.cmbFiadbTreeTable.Text.Trim().Length == 0 ||
				this.cmbFiadbSiteTreeTable.Text.Trim().Length==0)
			{
				MessageBox.Show("Enter a value for each table","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				return;
			}

			this.btnFilterNext.Enabled=true;
			this.btnFilterFinish.Enabled=false;
			
			
			this.rdoFilterByFile.Text = "Filter By File (Text File Containing Plot_CN numbers)";
		

			this.grpboxFilter.Visible=true;
			this.grpboxMDBFiadbInput.Visible=false;



		


		}
		private bool LoadMDBFiadbPopEvalTable()
		{
			string strCN="";
			string strCNDelimited="";
			string strEvalId="";
			string strEvalIdDelimited="";
			string strRsCd="";
			string strRsCdDelimited="";
			string strStateCd="";
			string strStateCdDelimited="";
			string strLocNm="";
			string strLocNmDelimited="";
			string strEvalDesc="";
			string strEvalDescDelimited="";
			string strRptYr="";
			string strRptYrDelimited="";
			string strNotes="";
			string strNotesDelimited="";
			int x=0;
			m_intError=0;
			bool bLoad=false;
			if (this.m_ado==null)
				this.m_ado = new ado_data_access();
			if (m_oDatasource==null) this.InitializeDatasource();
				
			try
			{
				//check if the eval list box has no values
				if (this.lstFIADBInv.Items.Count==0)
				{
					bLoad=true;
				}
				//see if the same values in the list as the table
				m_ado.SqlQueryReader(m_ado.getMDBConnString(this.txtMDBFiadbInputFile.Text.Trim(),"",""),"SELECT * FROM " + this.cmbFiadbPopEvalTable.Text.Trim());
				if (m_ado.m_intError==0)
				{
					if (m_ado.m_OleDbDataReader.HasRows)
					{
						while (m_ado.m_OleDbDataReader.Read())
						{
							//initialize eval values
							strCN="";
							strEvalId="";
							strRsCd="";
							strStateCd="";
							strLocNm="";
							strEvalDesc="";
							strRptYr="";
							strNotes="";
							strCN =m_ado.m_OleDbDataReader["cn"].ToString();
							strEvalId =m_ado.m_OleDbDataReader["evalid"].ToString();
							strRsCd= m_ado.m_OleDbDataReader["RsCd"].ToString();
							strStateCd = m_ado.m_OleDbDataReader["statecd"].ToString();
							if (m_ado.m_OleDbDataReader["location_nm"] != System.DBNull.Value)
								strLocNm = m_ado.m_OleDbDataReader["location_nm"].ToString();
							if (m_ado.m_OleDbDataReader["eval_descr"] != System.DBNull.Value)
								strEvalDesc = m_ado.m_OleDbDataReader["eval_descr"].ToString();
							if (m_ado.m_OleDbDataReader["report_year_nm"] != System.DBNull.Value)
								strRptYr = m_ado.m_OleDbDataReader["report_year_nm"].ToString();
							if (m_ado.m_OleDbDataReader["notes"] != System.DBNull.Value)
								strNotes = m_ado.m_OleDbDataReader["notes"].ToString();	
							//string all the eval records
							strCNDelimited=strCNDelimited + strCN + " " + "#";
							strEvalIdDelimited=strEvalIdDelimited + strEvalId + " " + "#";
							strRsCdDelimited = strRsCdDelimited + strRsCd + " " + "#";
							strStateCdDelimited = strStateCdDelimited + strStateCd + " " + "#";
							strLocNmDelimited = strLocNmDelimited + strLocNm + " " + "#";
							strEvalDescDelimited = strEvalDescDelimited + strEvalDesc + " " + "#";
							strRptYrDelimited = strRptYrDelimited + strRptYr + " " + "#";
							strNotesDelimited = strNotesDelimited + strNotes + " " + "#";
							if (!bLoad)
							{
								//see if the tables eval row is found in the list box
								for (x=0;x<=this.lstFIADBInv.Items.Count-1;x++)
								{
									if (strEvalId.Trim() == this.lstFIADBInv.Items[x].SubItems[0].Text.Trim() && 
										strRsCd.Trim() == this.lstFIADBInv.Items[x].SubItems[1].Text.Trim() && 
										strStateCd.Trim() == this.lstFIADBInv.Items[x].SubItems[2].Text.Trim() &&
										strLocNm.Trim() == this.lstFIADBInv.Items[x].SubItems[3].Text.Trim() &&
										strEvalDesc.Trim() == this.lstFIADBInv.Items[x].SubItems[4].Text.Trim() &&
										strRptYr.Trim() == this.lstFIADBInv.Items[x].SubItems[5].Text.Trim() && 
										strNotes.Trim() == this.lstFIADBInv.Items[x].SubItems[6].Text.Trim())
										break;
								}
								if (x > this.lstFIADBInv.Items.Count-1) 
								{
									//the eval table record is not found in the list box
									bLoad=true;
								}
							}
						}
						m_ado.m_OleDbDataReader.Close();
						while (m_ado.m_OleDbDataReader.IsClosed==false)
							System.Threading.Thread.Sleep(1000);
						
						if (bLoad)
						{
							//remove the delimiter from the end of the string list
							if (strCNDelimited.Trim().Length > 0) strCNDelimited = strCNDelimited.Substring(0,strCNDelimited.Length - 1);
							if (strEvalIdDelimited.Trim().Length > 0) strEvalIdDelimited = strEvalIdDelimited.Substring(0,strEvalIdDelimited.Length - 1);
							if (strRsCdDelimited.Trim().Length > 0) strRsCdDelimited = strRsCdDelimited.Substring(0,strRsCdDelimited.Length - 1);
							if (strStateCdDelimited.Trim().Length > 0) strStateCdDelimited=strStateCdDelimited.Substring(0,strStateCdDelimited.Length - 1);
							if (strLocNmDelimited.Trim().Length > 0) strLocNmDelimited=strLocNmDelimited.Substring(0,strLocNmDelimited.Length - 1);
							if (strEvalDescDelimited.Trim().Length > 0) strEvalDescDelimited=strEvalDescDelimited.Substring(0,strEvalDescDelimited.Length - 1);
							if (strRptYrDelimited.Trim().Length > 0) strRptYrDelimited=strRptYrDelimited.Substring(0,strRptYrDelimited.Length - 1);
							if (strNotesDelimited.Trim().Length > 0) strNotesDelimited=strNotesDelimited.Substring(0,strNotesDelimited.Length - 1);

							//create a temporary mdb file with links to all the project tables
							this.m_strTempMDBFile = m_oDatasource.CreateMDBAndTableDataSourceLinks();

							//get a connection string for the temp mdb file
							this.m_strTempMDBFileConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"","");
							this.m_ado.OpenConnection(m_strTempMDBFileConn);
							if (m_ado.m_intError==0)
							{

								//delete the current eval records that have a value of 9
								m_ado.m_strSQL = "DELETE FROM " + this.m_strPopEvalTable + " WHERE biosum_status_cd=9";
								m_ado.SqlNonQuery(m_ado.m_OleDbConnection,m_ado.m_strSQL);
								if (m_ado.m_intError==0)
								{
									//covert the string lists to arrays
									utils oUtils = new utils();
									string[] strCNArray = oUtils.ConvertListToArray(strCNDelimited,"#");
									string[] strEvalIdArray = oUtils.ConvertListToArray(strEvalIdDelimited,"#");
									string[] strRsCdArray = oUtils.ConvertListToArray(strRsCdDelimited,"#");
									string[] strStateCdArray = oUtils.ConvertListToArray(strStateCdDelimited,"#");
									string[] strLocNmArray = oUtils.ConvertListToArray(strLocNmDelimited,"#");
									string[] strEvalDescArray = oUtils.ConvertListToArray(strEvalDescDelimited,"#");
									string[] strRptYrArray = oUtils.ConvertListToArray(strRptYrDelimited,"#");
									string[] strNotesArray = oUtils.ConvertListToArray(strNotesDelimited,"#");
									oUtils=null;
									//insert the evaluation records into the biosum evaluation table
									for (x=0;x<=strEvalIdArray.Length-1;x++)
									{
										m_ado.m_strSQL = "DELETE FROM " + this.m_strPopEvalTable + " " + 
											             "WHERE TRIM(cn)='" +  strCNArray[x].Trim() + "' AND " + 
																"rscd=" + strRsCdArray[x].Trim() + " AND " + 
											                    "evalid=" + strEvalIdArray[x].Trim();
										m_ado.SqlNonQuery(m_ado.m_OleDbConnection,m_ado.m_strSQL);	
										if (m_ado.m_intError!=0)
										{
											this.m_intError=m_ado.m_intError;
											break;
										}
										m_ado.m_strSQL="INSERT INTO " + this.m_strPopEvalTable + " " + 
											"(CN,RSCD,EVALID,EVAL_DESCR,STATECD," + 
											"LOCATION_NM,REPORT_YEAR_NM,NOTES," + 
											"START_INVYR,END_INVYR,BIOSUM_STATUS_CD) VALUES " + 
											"('" + strCNArray[x].Trim() + "'," + 
											strRsCdArray[x].Trim() + "," + 
											strEvalIdArray[x].Trim() + ",'" + 
											strEvalDescArray[x] + "'," + 
											strStateCdArray[x] + ",'" + 
											strLocNmArray[x] + "','" + 
											strRptYrArray[x] + "','" + 
											strNotesArray[x] + "',null,null,9)";
										m_ado.SqlNonQuery(m_ado.m_OleDbConnection,m_ado.m_strSQL);	
										if (m_ado.m_intError!=0)
										{
											this.m_intError=m_ado.m_intError;
											break;
										}
									}
								}          							
								else m_intError=m_ado.m_intError;
								m_ado.m_OleDbConnection.Close();
								while (this.m_ado.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
									System.Threading.Thread.Sleep(1000);
							}
							else m_intError=m_ado.m_intError;
						}
					}
					else
					{
						MessageBox.Show("There are no Population Evaluations in the " + this.cmbFiadbPopEvalTable.Text + " table ",
							"FIA Biosum",
							System.Windows.Forms.MessageBoxButtons.OK,
							System.Windows.Forms.MessageBoxIcon.Exclamation);
						m_intError=-1;
						m_ado.m_OleDbDataReader.Close();
						while (m_ado.m_OleDbDataReader.IsClosed==false)
							System.Threading.Thread.Sleep(1000);
						bLoad=false;
						
					}
				}

			}
			catch (Exception e)
			{
				this.m_intError=-1;
				MessageBox.Show(e.Message,
					"FIA Biosum",
					System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				return false;
			}
			return bLoad;


		}

		private bool LoadMDBFiadbPopEvalTable2()
		{
			string strCN="";
			string strCNDelimited="";
			string strEvalId="";
			string strEvalIdDelimited="";
			string strRsCd="";
			string strRsCdDelimited="";
			string strStateCd="";
			string strStateCdDelimited="";
			string strLocNm="";
			string strLocNmDelimited="";
			string strEvalDesc="";
			string strEvalDescDelimited="";
			string strRptYr="";
			string strRptYrDelimited="";
			string strNotes="";
			string strNotesDelimited="";
			int x=0;
			m_intError=0;
			bool bLoad=false;
			if (this.m_ado==null)
				this.m_ado = new ado_data_access();
			if (m_oDatasource==null) this.InitializeDatasource();
				
			try
			{
				//check if the eval list box has no values
				if (this.lstFIADBInv.Items.Count==0)
				{
					bLoad=true;
				}
				//see if the same values in the list as the table
				m_ado.SqlQueryReader(m_ado.getMDBConnString(this.txtMDBFiadbInputFile.Text.Trim(),"",""),"SELECT * FROM " + this.cmbFiadbPopEvalTable.Text.Trim());
				if (m_ado.m_intError==0)
				{
					if (m_ado.m_OleDbDataReader.HasRows)
					{
						while (m_ado.m_OleDbDataReader.Read())
						{
							//initialize eval values
							strCN="";
							strEvalId="";
							strRsCd="";
							strStateCd="";
							strLocNm="";
							strEvalDesc="";
							strRptYr="";
							strNotes="";
							strCN =m_ado.m_OleDbDataReader["cn"].ToString();
							strEvalId =m_ado.m_OleDbDataReader["evalid"].ToString();
							strRsCd= m_ado.m_OleDbDataReader["RsCd"].ToString();
							strStateCd = m_ado.m_OleDbDataReader["statecd"].ToString();
							if (m_ado.m_OleDbDataReader["location_nm"] != System.DBNull.Value)
								strLocNm = m_ado.m_OleDbDataReader["location_nm"].ToString();
							if (m_ado.m_OleDbDataReader["eval_descr"] != System.DBNull.Value)
								strEvalDesc = m_ado.m_OleDbDataReader["eval_descr"].ToString();
							if (m_ado.m_OleDbDataReader["report_year_nm"] != System.DBNull.Value)
								strRptYr = m_ado.m_OleDbDataReader["report_year_nm"].ToString();
							if (m_ado.m_OleDbDataReader["notes"] != System.DBNull.Value)
								strNotes = m_ado.m_OleDbDataReader["notes"].ToString();	
							//string all the eval records
							strCNDelimited=strCNDelimited + strCN + " " + "#";
							strEvalIdDelimited=strEvalIdDelimited + strEvalId + " " + "#";
							strRsCdDelimited = strRsCdDelimited + strRsCd + " " + "#";
							strStateCdDelimited = strStateCdDelimited + strStateCd + " " + "#";
							strLocNmDelimited = strLocNmDelimited + strLocNm + " " + "#";
							strEvalDescDelimited = strEvalDescDelimited + strEvalDesc + " " + "#";
							strRptYrDelimited = strRptYrDelimited + strRptYr + " " + "#";
							strNotesDelimited = strNotesDelimited + strNotes + " " + "#";
							if (!bLoad)
							{
								//see if the tables eval row is found in the list box
								for (x=0;x<=this.lstFIADBInv.Items.Count-1;x++)
								{
									if (strEvalId.Trim() == this.lstFIADBInv.Items[x].SubItems[0].Text.Trim() && 
										strRsCd.Trim() == this.lstFIADBInv.Items[x].SubItems[1].Text.Trim() && 
										strStateCd.Trim() == this.lstFIADBInv.Items[x].SubItems[2].Text.Trim() &&
										strLocNm.Trim() == this.lstFIADBInv.Items[x].SubItems[3].Text.Trim() &&
										strEvalDesc.Trim() == this.lstFIADBInv.Items[x].SubItems[4].Text.Trim() &&
										strRptYr.Trim() == this.lstFIADBInv.Items[x].SubItems[5].Text.Trim() && 
										strNotes.Trim() == this.lstFIADBInv.Items[x].SubItems[6].Text.Trim())
										break;
								}
								if (x > this.lstFIADBInv.Items.Count) 
								{
									//the eval table record is not found in the list box
									bLoad=true;
								}
							}
						}
						m_ado.m_OleDbDataReader.Close();
						while (m_ado.m_OleDbDataReader.IsClosed==false)
							System.Threading.Thread.Sleep(1000);
						
						if (bLoad)
						{
							//remove the delimiter from the end of the string list
							if (strCNDelimited.Trim().Length > 0) strCNDelimited = strCNDelimited.Substring(0,strCNDelimited.Length - 1);
							if (strEvalIdDelimited.Trim().Length > 0) strEvalIdDelimited = strEvalIdDelimited.Substring(0,strEvalIdDelimited.Length - 1);
							if (strRsCdDelimited.Trim().Length > 0) strRsCdDelimited = strRsCdDelimited.Substring(0,strRsCdDelimited.Length - 1);
							if (strStateCdDelimited.Trim().Length > 0) strStateCdDelimited=strStateCdDelimited.Substring(0,strStateCdDelimited.Length - 1);
							if (strLocNmDelimited.Trim().Length > 0) strLocNmDelimited=strLocNmDelimited.Substring(0,strLocNmDelimited.Length - 1);
							if (strEvalDescDelimited.Trim().Length > 0) strEvalDescDelimited=strEvalDescDelimited.Substring(0,strEvalDescDelimited.Length - 1);
							if (strRptYrDelimited.Trim().Length > 0) strRptYrDelimited=strRptYrDelimited.Substring(0,strRptYrDelimited.Length - 1);
							if (strNotesDelimited.Trim().Length > 0) strNotesDelimited=strNotesDelimited.Substring(0,strNotesDelimited.Length - 1);

							//create a temporary mdb file with links to all the project tables
							this.m_strTempMDBFile = m_oDatasource.CreateMDBAndTableDataSourceLinks();

							//get a connection string for the temp mdb file
							this.m_strTempMDBFileConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"","");
							this.m_ado.OpenConnection(m_strTempMDBFileConn);
							if (m_ado.m_intError==0)
							{

								//delete the current eval records that have a value of 9
								m_ado.m_strSQL = "DELETE FROM " + this.m_strPopEvalTable + " WHERE biosum_status_cd=9";
								m_ado.SqlNonQuery(m_ado.m_OleDbConnection,m_ado.m_strSQL);
								if (m_ado.m_intError==0)
								{
									//covert the string lists to arrays
									utils oUtils = new utils();
									string[] strCNArray = oUtils.ConvertListToArray(strCNDelimited,"#");
									string[] strEvalIdArray = oUtils.ConvertListToArray(strEvalIdDelimited,"#");
									string[] strRsCdArray = oUtils.ConvertListToArray(strRsCdDelimited,"#");
									string[] strStateCdArray = oUtils.ConvertListToArray(strStateCdDelimited,"#");
									string[] strLocNmArray = oUtils.ConvertListToArray(strLocNmDelimited,"#");
									string[] strEvalDescArray = oUtils.ConvertListToArray(strEvalDescDelimited,"#");
									string[] strRptYrArray = oUtils.ConvertListToArray(strRptYrDelimited,"#");
									string[] strNotesArray = oUtils.ConvertListToArray(strNotesDelimited,"#");
									oUtils=null;
									//insert the evaluation records into the biosum evaluation table
									for (x=0;x<=strEvalIdArray.Length-1;x++)
									{
										m_ado.m_strSQL="INSERT INTO " + this.m_strPopEvalTable + " " + 
											"(CN,RSCD,EVALID,EVAL_DESCR,STATECD," + 
											"LOCATION_NM,REPORT_YEAR_NM,NOTES," + 
											"START_INVYR,END_INVYR,BIOSUM_STATUS_CD) VALUES " + 
											"('" + strCNArray[x].Trim() + "'," + 
											strRsCdArray[x].Trim() + "," + 
											strEvalIdArray[x].Trim() + ",'" + 
											strEvalDescArray[x] + "'," + 
											strStateCdArray[x] + ",'" + 
											strLocNmArray[x] + "','" + 
											strRptYrArray[x] + "','" + 
											strNotesArray[x] + "',null,null,9)";
										m_ado.SqlNonQuery(m_ado.m_OleDbConnection,m_ado.m_strSQL);	
										if (m_ado.m_intError!=0)
										{
											this.m_intError=m_ado.m_intError;
											break;
										}
									}
								}          							
								else m_intError=m_ado.m_intError;
								m_ado.m_OleDbConnection.Close();
								while (this.m_ado.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
									System.Threading.Thread.Sleep(1000);
							}
							else m_intError=m_ado.m_intError;
						}
					}
					else
					{
						MessageBox.Show("There are no Population Evaluations in the " + this.cmbFiadbPopEvalTable.Text + " table ",
							"FIA Biosum",
							System.Windows.Forms.MessageBoxButtons.OK,
							System.Windows.Forms.MessageBoxIcon.Exclamation);
						m_intError=-1;
						m_ado.m_OleDbDataReader.Close();
						while (m_ado.m_OleDbDataReader.IsClosed==false)
							System.Threading.Thread.Sleep(1000);
						bLoad=false;
						
					}
				}

			}
			catch (Exception e)
			{
				this.m_intError=-1;
				MessageBox.Show(e.Message,
					"FIA Biosum",
					System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				return false;
			}
			return bLoad;


		}
		private void LoadMDBFiadbPopFiles()
		{
				 this.m_strCurrentProcess="mdbFiadbInputPopTables";	
				this.m_ado = new ado_data_access();
				    
			

                    
				
				//create a temporary mdb file with links to all the project tables
				//and return the name of the file that contains the links
				this.m_strTempMDBFile = m_oDatasource.CreateMDBAndTableDataSourceLinks();

				this.m_bLoadStateCountyList=true;
				this.m_bLoadStateCountyPlotList=true;

				this.StartTherm("2","Add MS Access Pop Table Data");
				this.m_frmTherm.progressBar2.Maximum=3;
				this.m_frmTherm.progressBar2.Minimum=0;
				this.m_frmTherm.progressBar2.Value=0;
				this.m_frmTherm.lblMsg2.Text = "Overall Progress";
				this.m_strTableType="POPULATION STRATUM";
			    
				this.m_strCurrentFiadbTable = this.cmbFiadbPopStratumTable.Text;
			    this.m_strCurrentFiadbInputFile = this.txtMDBFiadbInputFile.Text;
			    this.m_strCurrentBiosumTable=this.m_strPopStratumTable;

			

				this.thdProcessRecords = new Thread(new ThreadStart(mdbFiadbInputPopTables));
				this.thdProcessRecords.IsBackground = true;
				this.thdProcessRecords.Start();
				while (thdProcessRecords.IsAlive)
				{
					thdProcessRecords.Join(1000);
					System.Windows.Forms.Application.DoEvents();

				}
				this.m_frmTherm.progressBar2.Value=1;
				thdProcessRecords=null;
				if (m_intError==0)
				{
					this.m_strLoadedPopStratumInputTable = this.m_strCurrentFiadbTable;
					this.m_strTableType="POPULATION ESTIMATION UNIT";
					this.m_frmTherm.lblMsg.Text = "pop estimation unit table";
					this.m_strCurrentFiadbTable = this.cmbFiadbPopEstUnitTable.Text;
					this.m_strCurrentFiadbInputFile = this.txtMDBFiadbInputFile.Text;
					this.m_strCurrentBiosumTable = this.m_strPopEstUnitTable;
					this.thdProcessRecords = new Thread(new ThreadStart(mdbFiadbInputPopTables));
					this.thdProcessRecords.IsBackground = true;
					this.thdProcessRecords.Start();
					while (thdProcessRecords.IsAlive)
					{
						thdProcessRecords.Join(1000);
						System.Windows.Forms.Application.DoEvents();

					}
					thdProcessRecords=null;

				}
				this.m_frmTherm.progressBar2.Value=2;
				if (m_intError==0)
				{
					this.m_strLoadedPopEstUnitInputTable = this.m_strCurrentFiadbTable;
					this.m_strTableType="POPULATION PLOT STRATUM ASSIGNMENT";
					this.m_frmTherm.lblMsg.Text = "ppsa table";
					this.m_strCurrentFiadbTable = this.cmbFiadbPpsaTable.Text;
					this.m_strCurrentBiosumTable = this.m_strPpsaTable;
					this.thdProcessRecords = new Thread(new ThreadStart(mdbFiadbInputPopTables));
					this.thdProcessRecords.IsBackground = true;
					this.thdProcessRecords.Start();
					while (thdProcessRecords!=null && thdProcessRecords.IsAlive)
					{
						thdProcessRecords.Join(1000);
						System.Windows.Forms.Application.DoEvents();

					}
					thdProcessRecords=null;

				}
				if (this.m_intError==0)
				{
					this.m_strLoadedPpsaInputTable=this.m_strCurrentFiadbTable;
					this.m_strLoadedFiadbInputFile=this.m_strCurrentFiadbInputFile;
				}
				this.m_frmTherm.progressBar2.Value=this.m_frmTherm.progressBar2.Maximum;
				System.Threading.Thread.Sleep(2000);
				this.m_frmTherm.Close();
				this.m_frmTherm = null;

			    this.m_strCurrentProcess="";	
			
		}
		private void mdbFiadbInputPopTables()
		{
		   
			string strFields="";
			
			
			int x=0;
			int y=0;
			string strCol="";
			
			this.m_intError=0;		
			
			string strSourceFile=this.m_strCurrentFiadbInputFile;
			string strSourceTable=this.m_strCurrentFiadbTable;
			string strDestTable=this.m_strCurrentBiosumTable;
			string strSourceTableLink="fiadb_input_" + strSourceTable;
			

			
                    
			try
			{
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", 4);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);

				//instatiate dao for creating links in the temp table
				//to the fiadb plot, cond, and tree input tables
				dao_data_access p_dao1 = new dao_data_access();

				//create links to the fiadb input tables in the temp mdb file
				p_dao1.CreateTableLink(this.m_strTempMDBFile,
					strSourceTableLink,
					strSourceFile.Trim(),
					strSourceTable.Trim());
		    
				//destroy the object and release it from memory
				p_dao1 = null;

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 1);


				//get an ado connection string for the temp mdb file
				this.m_strTempMDBFileConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"","");

   					
				//create a new connection to the temp MDB file
				this.m_connTempMDBFile = new System.Data.OleDb.OleDbConnection();

				//open the connection to the temp mdb file 
				this.m_ado.OpenConnection(this.m_strTempMDBFileConn,ref this.m_connTempMDBFile);

				this.m_ado.m_strSQL = "DELETE FROM " + strDestTable;
				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 2);

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				//get the fiabiosum table structures
				System.Data.DataTable dtDestSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + strDestTable);
				//get the fiadb table structures
				System.Data.DataTable dtSourceSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + strSourceTableLink);

				
				//build field list string to insert sql by matching 
				//up the column names in the biosum plot table and the fiadb plot table
				strFields = "";
				for (x=0; x<=dtDestSchema.Rows.Count-1;x++)
				{
					strCol = dtDestSchema.Rows[x]["columnname"].ToString().Trim();
					//see if there is an equivalent FIADB column
					for (y=0; y<=dtSourceSchema.Rows.Count-1;y++)
					{
						if (strCol.Trim().ToUpper() == dtSourceSchema.Rows[y]["columnname"].ToString().ToUpper())
						{
							if (strFields.Trim().Length == 0)
							{
								strFields = strCol;
							}
							else
							{	
								strFields += "," + strCol;
							}
							break;
						}
					}
				}

				this.m_ado.m_strSQL = "INSERT INTO " + strDestTable + " (" + strFields + ")" + 
					" SELECT " + strFields + " FROM " + strSourceTableLink + "   " + 
					"WHERE rscd = " + this.m_strCurrFIADBRsCd + " AND " + 
						  "evalid = " + this.m_strCurrFIADBEvalId;
				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 3);
				this.m_ado.m_strSQL = "UPDATE " + strDestTable + " d INNER JOIN " + strSourceTableLink + " s ON d.cn=s.cn " +
					"SET d.biosum_status_cd=9";
				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 4);
				this.m_connTempMDBFile.Close();
				while (this.m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
					System.Threading.Thread.Sleep(1000);
				this.m_connTempMDBFile=null;
				
			}
			catch 
			{
				this.m_intError=-1;
				if (this.m_connTempMDBFile != null)
				{
					if (this.m_connTempMDBFile.State == System.Data.ConnectionState.Open)
					{
						this.m_connTempMDBFile.Close();
						while (this.m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
							System.Threading.Thread.Sleep(1000);
						this.m_connTempMDBFile=null;
					}
				}
				((frmDialog)this.ParentForm).Enabled=true;
			}
			finally
			{
				if (this.m_connTempMDBFile != null)
				{
					if (this.m_connTempMDBFile.State == System.Data.ConnectionState.Open)
					{
						this.m_connTempMDBFile.Close();
						while (this.m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
							System.Threading.Thread.Sleep(1000);
						this.m_connTempMDBFile=null;
					}
				}
				((frmDialog)this.ParentForm).Enabled=true;
				
				this.AddPlotRecordsFinished();
			}
			((frmDialog)this.ParentForm).Enabled=true;
			
			this.AddPlotRecordsFinished();

		}

		private void btnMDBFiadbInputCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}

		private void lstFIADBInv_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (lstFIADBInv.SelectedItems.Count==0) return;
			this.m_strCurrFIADBEvalId = this.lstFIADBInv.SelectedItems[0].Text.Trim();
			this.m_strCurrFIADBRsCd = this.lstFIADBInv.SelectedItems[0].SubItems[1].Text.Trim();
		
		}

		private void btnMDBSiteTreeBrowse_Click(object sender, System.EventArgs e)
		{
			this.GetMDBFileAndTable("MS Access Data File Containing Site Tree Table Data",
									"Select Site Tree Table",
									ref this.txtMDBSiteTree,
									ref this.txtMDBSiteTreeTable);
		}
		private void GetMDBFileAndTable(string p_strDialogTitleGetMDBFile,
			string p_strDialogTitleGetMDBTable,
			ref System.Windows.Forms.TextBox p_txtMDBFile,
			ref System.Windows.Forms.TextBox p_txtMDBTable)
		{
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = p_strDialogTitleGetMDBFile;
			OpenFileDialog1.Filter = "Microsoft Access Database File (*.MDB,*.MDE,*.ACCDB) |*.mdb;*.mde;*.accdb";
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK) 
			{
				if (OpenFileDialog1.FileName.Trim().Length > 0) 
				{
					string strFullPath = OpenFileDialog1.FileName.Trim();
					if (strFullPath.Length > 0) 
					{
						utils p_utils = new utils();
						string strFile = p_utils.getFileName(strFullPath);
						string strDir = p_utils.getDirectory(strFullPath);
						p_utils = null;
						dao_data_access tempDao = new dao_data_access();
						tempDao.OpenDb(strFullPath);
						if (tempDao.m_intError == 0) 
						{
							frmDialog frmTemp = new frmDialog();
							frmTemp.Text = p_strDialogTitleGetMDBTable;
							frmTemp.uc_select_list_item1.lblMsg.Text= "Table contents of " + strFullPath;
							frmTemp.uc_select_list_item1.lblMsg.Visible = true;
							string strLargestString = frmTemp.uc_select_list_item1.lblMsg.Text;
						
							frmTemp.uc_project1.Visible=false;
							frmTemp.uc_select_list_item1.listBox1.Items.Clear();
							for (int x=0; x <= tempDao.m_DaoDatabase.TableDefs.Count - 1; x++)
							{
								if (tempDao.m_DaoDatabase.TableDefs[x].Name.IndexOf("MSys",0) < 0) 		
								{
									frmTemp.uc_select_list_item1.listBox1.Items.Add(tempDao.m_DaoDatabase.TableDefs[x].Name);
									if (tempDao.m_DaoDatabase.TableDefs[x].Name.Trim().Length > 
										strLargestString.Trim().Length)
										strLargestString = tempDao.m_DaoDatabase.TableDefs[x].Name;
								}

							}
                        
						
							tempDao.m_DaoDatabase.Close();
							tempDao.m_DaoDatabase = null;
						
							frmTemp.uc_select_list_item1.Initialize_Width(strLargestString);
							frmTemp.uc_select_list_item1.Visible=true;
							result = frmTemp.ShowDialog(this);
                        
						
							if (result == DialogResult.OK) 
							{
							
								p_txtMDBFile.Text = strFullPath;
								p_txtMDBTable.Text = frmTemp.uc_select_list_item1.listBox1.Text;
							}
					
							frmTemp.Close();
							frmTemp = null;
						}
						tempDao = null;
					}
				}
			}
			else 
			{
			}
			OpenFileDialog1 = null;

		}

		private void btnSiteTreeBrowse_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "FIADB Site Tree Table Data";
			OpenFileDialog1.Filter = "Comma Delimited Text File (*.CSV;*.TXT;*.DAT) |*.csv;*.txt;*.dat";
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK) 
			{
				if (OpenFileDialog1.FileName.Trim().Length > 0) 
				{
					this.m_strSiteTreeTxtInputFile = OpenFileDialog1.FileName.Trim();
					this.txtSiteTree.Text = this.m_strSiteTreeTxtInputFile;
				}
			}
			else 
			{
			}
			OpenFileDialog1 = null;
		
		}

        private void cmbCondPropPercent_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }
        /*
        public class FIADB_Adjustments
        {
            public FIADB_Adjustments()
            {
                if (!System.IO.File.Exists(
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\DB\\BIOSUM_RECALC_FIADB_ADJUSTMENTS.ACCDB"))
                {
                    dao_data_access oDao = new dao_data_access();
                    oDao.CreateMDB(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\DB\\biosum_recalc_fiadb_adjustments.accdb");
                    oDao.m_DaoWorkspace.Close();
                    oDao.m_DaoWorkspace = null;
                    oDao = null;
                }
            }
            public enum ModeValues
            {
                ADD,
                DELETE,
            }
            private ModeValues _EditMode=ModeValues.ADD;
            public ModeValues EditMode
            {
                get { return _EditMode; }
                set { _EditMode = value; }
            }
            private string _strMSAccessDbFile = "";
            public string MSAccessDbFile
            {
                get { return _strMSAccessDbFile; }
                set { _strMSAccessDbFile = value; }
            }
           
            private void ImportCSVFiles(
                string p_strPlotFile,
                string p_strPopEstUnitFile,
                string p_strPopEvalFile,
                string p_strPopStratumFile,
                string p_strPPSAFile,
                string p_strCondFile)
            {

            }
            public void FIADB_Adjustments_Process()
            {

            }

         
        }
         */
	}
    
	
}
