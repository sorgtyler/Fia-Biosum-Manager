using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_tree_diam_groups_list.
	/// </summary>
	public class uc_tree_diam_groups_list : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnEdit;
		private System.Windows.Forms.Button btnNew;
		public int m_intDialogHt;
		public int m_intDialogWd;
		private System.Windows.Forms.ListView lstTreeDiam;
		private int m_intError;
        private string m_strError;
		private System.Windows.Forms.Button btnDefault;
		public System.Windows.Forms.Button btnSave;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnClear;
		private System.Windows.Forms.Button btnHelp;
		private FIA_Biosum_Manager.ado_data_access m_ado;
		private System.Windows.Forms.Button btnDelete;

        // scenario-specific variables
        private string _strScenarioId = "";
        private frmProcessorScenario _frmProcessorScenario = null;


		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_tree_diam_groups_list()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.m_intDialogHt = this.groupBox1.Top + this.btnClose.Top + this.btnClose.Height + 20;
			this.m_intDialogWd = this.groupBox1.Left + this.btnClose.Left + this.btnClose.Width + 20;

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

        public frmProcessorScenario ReferenceProcessorScenarioForm
        {
            get { return this._frmProcessorScenario; }
            set { this._frmProcessorScenario = value; }
        }
        public string ScenarioId
        {
            get { return _strScenarioId; }
            set { _strScenarioId = value; }
        }


		public void loadvalues()
		{
            string strId="";
			string strMin="";
			string strMax="";
			string strDef="";
			      
			this.lstTreeDiam.Clear();
			this.lstTreeDiam.Columns.Add("Group ID", 60, HorizontalAlignment.Left);
			this.lstTreeDiam.Columns.Add("Minimum Diameter", 150, HorizontalAlignment.Left);
			this.lstTreeDiam.Columns.Add("Maximum Diameter", 150, HorizontalAlignment.Left);
			this.lstTreeDiam.Columns.Add("Definition", 200, HorizontalAlignment.Left);

			this.m_intError=0;

            ScenarioId = this.ReferenceProcessorScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
            string strDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsDbFile;
            ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.LoadTreeDiameterGroupValues(strDbFile,
                ScenarioId, ReferenceProcessorScenarioForm.m_oProcessorScenarioItem);

					try
					{
						//load up each row in the FIADB plot input table
                        int x;
                        for (x = 0; x <= ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Count - 1; x++)
                        {
                            ProcessorScenarioItem.TreeDiamGroupsItem p_oTreeGroupsItem = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Item(x);
							
                            strId="";
							strMax="";
							strMin="";
							strDef="";
							
                            //make sure the row is not null values
							if (p_oTreeGroupsItem.DiamClass.Trim().Length > 0)
							{
                                strId = p_oTreeGroupsItem.DiamGroup;
								strMin = p_oTreeGroupsItem.MinDiam;
                                strMax = p_oTreeGroupsItem.MaxDiam;
                                strDef = p_oTreeGroupsItem.DiamClass;
								this.lstTreeDiam.BeginUpdate();
								System.Windows.Forms.ListViewItem listItem = new ListViewItem();
								listItem.Text=strId;
								listItem.SubItems.Add(strMin);
								listItem.SubItems.Add(strMax);
								listItem.SubItems.Add(strDef);
								this.lstTreeDiam.Items.Add(listItem);
								this.lstTreeDiam.EndUpdate();
								
							}

						}
						if (this.lstTreeDiam.Items.Count > 0)
						{                                                       
							if (this.lstTreeDiam.SelectedItems.Count == 0)
							{
								this.lstTreeDiam.Items[this.lstTreeDiam.Items.Count-1].Selected=true;
							}
						}
						((frmDialog)this.ParentForm).Enabled=true;
					}
					catch (Exception caught)
					{
						this.m_intError=-1;
						MessageBox.Show(caught.Message);
					}
		}

        public void loadvalues_FromProperties()
        {
            string strId = "";
            string strMin = "";
            string strMax = "";
            string strDef = "";

            this.lstTreeDiam.Clear();
            this.lstTreeDiam.Columns.Add("Group ID", 60, HorizontalAlignment.Left);
            this.lstTreeDiam.Columns.Add("Minimum Diameter", 150, HorizontalAlignment.Left);
            this.lstTreeDiam.Columns.Add("Maximum Diameter", 150, HorizontalAlignment.Left);
            this.lstTreeDiam.Columns.Add("Definition", 200, HorizontalAlignment.Left);

            if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection != null)
            {
                try
                {
                    //load up each row from SCENARIO_TREE_DIAM_GROUPS table
                    int x;
                    for (x = 0; x <= ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Count - 1; x++)
                    {
                        ProcessorScenarioItem.TreeDiamGroupsItem p_oTreeGroupsItem = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Item(x);

                        strId = "";
                        strMax = "";
                        strMin = "";
                        strDef = "";

                        //make sure the row is not null values
                        if (p_oTreeGroupsItem.DiamClass.Trim().Length > 0)
                        {
                            strId = p_oTreeGroupsItem.DiamGroup;
                            strMin = p_oTreeGroupsItem.MinDiam;
                            strMax = p_oTreeGroupsItem.MaxDiam;
                            strDef = p_oTreeGroupsItem.DiamClass;
                            this.lstTreeDiam.BeginUpdate();
                            System.Windows.Forms.ListViewItem listItem = new ListViewItem();
                            listItem.Text = strId;
                            listItem.SubItems.Add(strMin);
                            listItem.SubItems.Add(strMax);
                            listItem.SubItems.Add(strDef);
                            this.lstTreeDiam.Items.Add(listItem);
                            this.lstTreeDiam.EndUpdate();
                        }

                    }
                    if (this.lstTreeDiam.Items.Count > 0)
                    {
                        if (this.lstTreeDiam.SelectedItems.Count == 0)
                        {
                            this.lstTreeDiam.Items[this.lstTreeDiam.Items.Count - 1].Selected = true;
                        }
                    }
                    ((frmDialog)this.ParentForm).Enabled = true;
                }
                catch (Exception caught)
                {
                    this.m_intError = -1;
                    MessageBox.Show(caught.Message);
                }
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
			this.btnHelp = new System.Windows.Forms.Button();
			this.btnClear = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnSave = new System.Windows.Forms.Button();
			this.btnDefault = new System.Windows.Forms.Button();
			this.btnNew = new System.Windows.Forms.Button();
			this.btnEdit = new System.Windows.Forms.Button();
			this.btnClose = new System.Windows.Forms.Button();
			this.lstTreeDiam = new System.Windows.Forms.ListView();
			this.lblTitle = new System.Windows.Forms.Label();
			this.btnDelete = new System.Windows.Forms.Button();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btnDelete);
			this.groupBox1.Controls.Add(this.btnHelp);
			this.groupBox1.Controls.Add(this.btnClear);
			this.groupBox1.Controls.Add(this.btnCancel);
			this.groupBox1.Controls.Add(this.btnSave);
			this.groupBox1.Controls.Add(this.btnDefault);
			this.groupBox1.Controls.Add(this.btnNew);
			this.groupBox1.Controls.Add(this.btnEdit);
			this.groupBox1.Controls.Add(this.btnClose);
			this.groupBox1.Controls.Add(this.lstTreeDiam);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(568, 304);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// btnHelp
			// 
			this.btnHelp.Location = new System.Drawing.Point(8, 259);
			this.btnHelp.Name = "btnHelp";
			this.btnHelp.Size = new System.Drawing.Size(64, 32);
			this.btnHelp.TabIndex = 8;
			this.btnHelp.Text = "Help";
			// 
			// btnClear
			// 
			this.btnClear.Location = new System.Drawing.Point(350, 216);
			this.btnClear.Name = "btnClear";
			this.btnClear.Size = new System.Drawing.Size(64, 32);
			this.btnClear.TabIndex = 7;
			this.btnClear.Text = "Clear All";
			this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
			// 
			// btnCancel
			// 
			this.btnCancel.Location = new System.Drawing.Point(478, 216);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(64, 32);
			this.btnCancel.TabIndex = 6;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// btnSave
			// 
			this.btnSave.Enabled = false;
			this.btnSave.Location = new System.Drawing.Point(414, 216);
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(64, 32);
			this.btnSave.TabIndex = 5;
			this.btnSave.Text = "Save";
			this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
			// 
			// btnDefault
			// 
			this.btnDefault.Location = new System.Drawing.Point(28, 216);
			this.btnDefault.Name = "btnDefault";
			this.btnDefault.Size = new System.Drawing.Size(114, 32);
			this.btnDefault.TabIndex = 2;
			this.btnDefault.Text = "Use Default Values";
			this.btnDefault.Click += new System.EventHandler(this.btnDefault_Click);
			// 
			// btnNew
			// 
			this.btnNew.Location = new System.Drawing.Point(158, 216);
			this.btnNew.Name = "btnNew";
			this.btnNew.Size = new System.Drawing.Size(64, 32);
			this.btnNew.TabIndex = 3;
			this.btnNew.Text = "New";
			this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
			// 
			// btnEdit
			// 
			this.btnEdit.Location = new System.Drawing.Point(222, 216);
			this.btnEdit.Name = "btnEdit";
			this.btnEdit.Size = new System.Drawing.Size(64, 32);
			this.btnEdit.TabIndex = 4;
			this.btnEdit.Text = "Edit";
			this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
			// 
			// btnClose
			// 
			this.btnClose.Location = new System.Drawing.Point(462, 259);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(96, 32);
			this.btnClose.TabIndex = 9;
			this.btnClose.Text = "Close";
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// lstTreeDiam
			// 
			this.lstTreeDiam.FullRowSelect = true;
			this.lstTreeDiam.GridLines = true;
			this.lstTreeDiam.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
			this.lstTreeDiam.HideSelection = false;
			this.lstTreeDiam.Location = new System.Drawing.Point(16, 48);
			this.lstTreeDiam.MultiSelect = false;
			this.lstTreeDiam.Name = "lstTreeDiam";
			this.lstTreeDiam.Size = new System.Drawing.Size(536, 160);
			this.lstTreeDiam.TabIndex = 1;
			this.lstTreeDiam.View = System.Windows.Forms.View.Details;
			this.lstTreeDiam.SelectedIndexChanged += new System.EventHandler(this.lstTreeDiam_SelectedIndexChanged);
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(562, 24);
			this.lblTitle.TabIndex = 0;
			this.lblTitle.Text = "Tree Diameter Groups List";
			// 
			// btnDelete
			// 
			this.btnDelete.Enabled = false;
			this.btnDelete.Location = new System.Drawing.Point(286, 216);
			this.btnDelete.Name = "btnDelete";
			this.btnDelete.Size = new System.Drawing.Size(64, 32);
			this.btnDelete.TabIndex = 10;
			this.btnDelete.Text = "Delete";
			this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
			// 
			// uc_tree_diam_groups_list
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_tree_diam_groups_list";
			this.Size = new System.Drawing.Size(568, 304);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

        private void btnClose_Click(object sender, System.EventArgs e)
        {
            if (this.btnSave.Enabled == true)
            {
                DialogResult result = MessageBox.Show("Save Changes Y/N", "Tree Diameter Groups", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    this.savevalues();
                }
            }
            this.ParentForm.Close();
            this.ParentForm.Dispose();
        }

		private void btnNew_Click(object sender, System.EventArgs e)
		{
			string strMax="";
			string strMin="";
			string strId = "";
            double dblMin = 0;

			//get the last id if there is one
			if (this.lstTreeDiam.Items.Count > 0)
			{
				strId = Convert.ToString(Convert.ToInt16(this.lstTreeDiam.Items[this.lstTreeDiam.Items.Count-1].Text) + 1);
				strMin = this.lstTreeDiam.Items[this.lstTreeDiam.Items.Count-1].SubItems[2].Text;
				if (strMin.IndexOf(".") < 0) strMin += ".0";

                dblMin = Convert.ToDouble(strMin) + 0.1;
                strMin = Convert.ToString(dblMin);
				
				strMax = strMin;
				//if (strMin.IndexOf(".") < 0) strMin += ".0";

			}
			else
			{
				strMax="0.0";
				strMin="0.0";
				strId = "1";
			}

			FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
		    frmTemp.MaximizeBox = false;
			frmTemp.BackColor = System.Drawing.SystemColors.Control;
			frmTemp.Text = "Database: Tree Diameter Groups (New)";
			frmTemp.Initialize_Plot_Tree_Diam_Edit_User_Control();

			frmTemp.uc_tree_diam_groups_edit1.txtGroupId = strId;
			frmTemp.uc_tree_diam_groups_edit1.txtMaximumDiameter = strMax;
			frmTemp.uc_tree_diam_groups_edit1.txtMinimumDiameter = strMin;


			//frmTemp.MdiParent = this;
			//frmTemp.Initialize_Plot_Tree_Diam_User_Control();


			frmTemp.Height=0;
			frmTemp.Width=0;
			if (frmTemp.uc_tree_diam_groups_edit1.Top + frmTemp.uc_tree_diam_groups_edit1.Height > frmTemp.ClientSize.Height + 2)
			{
				for (int x=1;;x++)
				{
					frmTemp.Height = x;
					if (frmTemp.uc_tree_diam_groups_edit1.Top + 
						frmTemp.uc_tree_diam_groups_edit1.Height < 
						frmTemp.ClientSize.Height)
					{
						break;
					}
				}

			}
			if (frmTemp.uc_tree_diam_groups_edit1.Left + frmTemp.uc_tree_diam_groups_edit1.Width > frmTemp.ClientSize.Width + 2)
			{
				for (int x=1;;x++)
				{
					frmTemp.Width = x;
					if (frmTemp.uc_tree_diam_groups_edit1.Left + 
						frmTemp.uc_tree_diam_groups_edit1.Width < 
						frmTemp.ClientSize.Width)
					{
						break;
					}
				}

			}
			frmTemp.Left = 0;
			frmTemp.Top = 0;
			frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
			if (result==System.Windows.Forms.DialogResult.OK)
			{
				strMax = frmTemp.uc_tree_diam_groups_edit1.txtMaximumDiameter.Trim();
				//remove the period if it is the last character
				if (strMax.IndexOf(".") == strMax.Length-1)
				{
					strMax = strMax.Replace(".","");
				}
				strMin = frmTemp.uc_tree_diam_groups_edit1.txtMinimumDiameter.Trim();
				//remove the period if it is the last character
				if (strMin.IndexOf(".") == strMin.Length-1)
				{
					strMin = strMin.Replace(".","");
				}
				this.lstTreeDiam.BeginUpdate();
				System.Windows.Forms.ListViewItem listItem = new ListViewItem();
				listItem.Text=frmTemp.uc_tree_diam_groups_edit1.txtGroupId;
				listItem.SubItems.Add(strMin);
				listItem.SubItems.Add(strMax);
				listItem.SubItems.Add(strMin + " - " + strMax);
				this.lstTreeDiam.Items.Add(listItem);
				this.lstTreeDiam.EndUpdate();
				if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			}
			frmTemp.Close();
			frmTemp.Dispose();
            frmTemp=null;
		}

		private void btnDefault_Click(object sender, System.EventArgs e)
		{
			this.lstTreeDiam.Clear();
			this.lstTreeDiam.Columns.Add("Group ID", 60, HorizontalAlignment.Left);
			this.lstTreeDiam.Columns.Add("Minimum Diameter", 150, HorizontalAlignment.Left);
			this.lstTreeDiam.Columns.Add("Maximum Diameter", 150, HorizontalAlignment.Left);
			this.lstTreeDiam.Columns.Add("Definition", 200, HorizontalAlignment.Left);

			this.m_intError=0;
			this.lstTreeDiam.BeginUpdate();
			this.lstTreeDiam.Items.Add("1");
			this.lstTreeDiam.Items[0].SubItems.Add("1");
			this.lstTreeDiam.Items[0].SubItems.Add("7");
			this.lstTreeDiam.Items[0].SubItems.Add("1 - 7");

			this.lstTreeDiam.Items.Add("2");
			this.lstTreeDiam.Items[1].SubItems.Add("7.1");
			this.lstTreeDiam.Items[1].SubItems.Add("10");
			this.lstTreeDiam.Items[1].SubItems.Add("7.1 - 10");

			this.lstTreeDiam.Items.Add("3");
			this.lstTreeDiam.Items[2].SubItems.Add("10.1");
			this.lstTreeDiam.Items[2].SubItems.Add("16");
			this.lstTreeDiam.Items[2].SubItems.Add("10.1 - 16");

			this.lstTreeDiam.Items.Add("4");
			this.lstTreeDiam.Items[3].SubItems.Add("16.1");
			this.lstTreeDiam.Items[3].SubItems.Add("21");
			this.lstTreeDiam.Items[3].SubItems.Add("16.1 - 21");

			this.lstTreeDiam.Items.Add("5");
			this.lstTreeDiam.Items[4].SubItems.Add("21.1");
			this.lstTreeDiam.Items[4].SubItems.Add("999");
			this.lstTreeDiam.Items[4].SubItems.Add("21.1 - 999");




			this.lstTreeDiam.EndUpdate();
			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;


		}

		private void btnClear_Click(object sender, System.EventArgs e)
		{
			if (this.lstTreeDiam.Items.Count > 0) this.btnSave.Enabled=true;
			this.lstTreeDiam.Clear();
			this.lstTreeDiam.Columns.Add("Group ID", 60, HorizontalAlignment.Left);
			this.lstTreeDiam.Columns.Add("Minimum Diameter", 150, HorizontalAlignment.Left);
			this.lstTreeDiam.Columns.Add("Maximum Diameter", 150, HorizontalAlignment.Left);
			this.lstTreeDiam.Columns.Add("Definition", 200, HorizontalAlignment.Left);
		}

		private void btnEdit_Click(object sender, System.EventArgs e)
		{
			string strMax="";
			string strMin="";
			string strId = "";

			if (this.lstTreeDiam.SelectedItems.Count == 0)
				return;



			
			strId = this.lstTreeDiam.SelectedItems[0].Text;
			strMin = this.lstTreeDiam.SelectedItems[0].SubItems[1].Text;
            strMax = this.lstTreeDiam.SelectedItems[0].SubItems[2].Text;
  		    if (strMin.IndexOf(".") < 0) strMin += ".0";
			if (strMax.IndexOf(".") < 0) strMax += ".0";

			FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
			frmTemp.MaximizeBox = false;
			frmTemp.BackColor = System.Drawing.SystemColors.Control;
			frmTemp.Text = "Database: Tree Diameter Groups (Edit)";
			frmTemp.Initialize_Plot_Tree_Diam_Edit_User_Control();

			frmTemp.uc_tree_diam_groups_edit1.txtGroupId = strId;
			frmTemp.uc_tree_diam_groups_edit1.txtMaximumDiameter = strMax;
			frmTemp.uc_tree_diam_groups_edit1.txtMinimumDiameter = strMin;


			frmTemp.Height=0;
			frmTemp.Width=0;
			if (frmTemp.uc_tree_diam_groups_edit1.Top + frmTemp.uc_tree_diam_groups_edit1.Height > frmTemp.ClientSize.Height + 2)
			{
				for (int x=1;;x++)
				{
					frmTemp.Height = x;
					if (frmTemp.uc_tree_diam_groups_edit1.Top + 
						frmTemp.uc_tree_diam_groups_edit1.Height < 
						frmTemp.ClientSize.Height)
					{
						break;
					}
				}

			}
			if (frmTemp.uc_tree_diam_groups_edit1.Left + frmTemp.uc_tree_diam_groups_edit1.Width > frmTemp.ClientSize.Width + 2)
			{
				for (int x=1;;x++)
				{
					frmTemp.Width = x;
					if (frmTemp.uc_tree_diam_groups_edit1.Left + 
						frmTemp.uc_tree_diam_groups_edit1.Width < 
						frmTemp.ClientSize.Width)
					{
						break;
					}
				}

			}
			frmTemp.Left = 0;
			frmTemp.Top = 0;
			frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
			if (result==System.Windows.Forms.DialogResult.OK)
			{
				strMax = frmTemp.uc_tree_diam_groups_edit1.txtMaximumDiameter.Trim();
				//remove the period if it is the last character
				if (strMax.IndexOf(".") == strMax.Length-1)
				{
					strMax = strMax.Replace(".","");
				}
				strMin = frmTemp.uc_tree_diam_groups_edit1.txtMinimumDiameter.Trim();
				//remove the period if it is the last character
				if (strMin.IndexOf(".") == strMin.Length-1)
				{
					strMin = strMin.Replace(".","");
				}
				this.lstTreeDiam.BeginUpdate();
				this.lstTreeDiam.SelectedItems[0].SubItems[1].Text = strMin;
				this.lstTreeDiam.SelectedItems[0].SubItems[2].Text = strMax;
				this.lstTreeDiam.SelectedItems[0].SubItems[3].Text = strMin + " - " + strMax;
				this.lstTreeDiam.EndUpdate();
				if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			}
			frmTemp.Close();
			frmTemp.Dispose();
			frmTemp=null;

		}

		private void btnSave_Click(object sender, System.EventArgs e)
		{
		  this.savevalues();
		}
		private void val_data()
		{
			int x;
			//int y;
			double dblMin;
            double dblMax;
			string strId1;
			string strId2;
            this.m_intError=0;

			if (this.lstTreeDiam.Items.Count==0)
			{
				MessageBox.Show("No tree diameter groups to save","Tree Diameter Groups",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}

			for (x=0;x<=this.lstTreeDiam.Items.Count-2;x++)
			{
				//make sure the preceding max is less than the next minimum
			   dblMax=Convert.ToDouble(this.lstTreeDiam.Items[x].SubItems[2].Text);
			   dblMin=Convert.ToDouble(this.lstTreeDiam.Items[x+1].SubItems[1].Text);
				if (dblMin <= dblMax)
				{ 
					strId1 = this.lstTreeDiam.Items[x].Text;
                    strId2 = this.lstTreeDiam.Items[x+1].Text ;
					MessageBox.Show("Group ID " + strId2 + " minimum diameter must be greater than Group ID " + strId1 + " maximum diameter","Tree Diameter Groups",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					this.m_intError=-1;
					return;
				}
			}            
		}
		public void savevalues()
		{


			val_data();
			if (this.m_intError==0)
			{
				//string strSQL;
				int x;
                string strMin;
				string strMax;
				string strDef;
				string strId;

                //
                //OPEN CONNECTION TO DB FILE CONTAINING PROCESSOR SCENARIO TABLES
                //
                //scenario mdb connection
                string strScenarioMDB =
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\db\\scenario_processor_rule_definitions.mdb";
                ado_data_access oAdo = new ado_data_access();
                oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioMDB, "", ""));
                if (oAdo.m_intError != 0)
                {
                    m_intError = m_ado.m_intError;
                    m_strError = m_ado.m_strError;
                    oAdo = null;
                    return;
                }

				if (this.m_intError==0)
				{
					//delete the current records
                    oAdo.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName +
                        " WHERE TRIM(UCASE(scenario_id)) = '" + ScenarioId.Trim().ToUpper() + "'";
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

                    if (oAdo.m_intError == 0)
					{
						for (x=0;x<=this.lstTreeDiam.Items.Count-1;x++)
						{
							strId = this.lstTreeDiam.Items[x].Text;
							strMin = this.lstTreeDiam.Items[x].SubItems[1].Text;
							strMax = this.lstTreeDiam.Items[x].SubItems[2].Text;
							strDef = this.lstTreeDiam.Items[x].SubItems[3].Text;

                            oAdo.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName + " " + 
								"(diam_group,diam_class,min_diam,max_diam,scenario_id) VALUES " + 
								"(" + strId + ",'" + strDef.Trim() + "'," + 
								strMin + "," + strMax + ",'" + ScenarioId + "');";
                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                            if (oAdo.m_intError != 0)
							{
								break;
							}
						}
					}
                    if (this.m_intError == 0 && oAdo.m_intError == 0)
					{
                        this.btnSave.Enabled=false;
					}

				}
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo = null;
			}

		}

        public void copyvalues(string strSourceScenario, string strTargetScenario)
        {
            m_ado = new ado_data_access();
            string strDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsDbFile;
            m_ado.OpenConnection(m_ado.getMDBConnString(strDbFile, "", ""));
            ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.LoadTreeDiameterGroupValues(m_ado,
                m_ado.m_OleDbConnection, strSourceScenario, ReferenceProcessorScenarioForm.m_oProcessorScenarioItem);

            if (this.m_intError == 0)
            {
                //delete the current records
                this.m_ado.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName +
                    " WHERE TRIM(UCASE(scenario_id)) = '" + strTargetScenario.Trim().ToUpper() + "'";
                this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection, this.m_ado.m_strSQL);

                if (this.m_ado.m_intError == 0)
                {
                    string strMin;
				    string strMax;
				    string strDef;
				    string strId;
                    for (int x = 0; x <= ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Count - 1; x++)

                    {
                        FIA_Biosum_Manager.ProcessorScenarioItem.TreeDiamGroupsItem oItem = 
                            ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Item(x);
                        strId = oItem.DiamClass;
                        strMin = this.lstTreeDiam.Items[x].SubItems[1].Text;
                        strMax = this.lstTreeDiam.Items[x].SubItems[2].Text;
                        strDef = this.lstTreeDiam.Items[x].SubItems[3].Text;

                        this.m_ado.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName + " " +
                            "(diam_group,diam_class,min_diam,max_diam,scenario_id) VALUES " +
                            "(" + strId + ",'" + strDef.Trim() + "'," +
                            strMin + "," + strMax + ",'" + ScenarioId + "');";
                        this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection, this.m_ado.m_strSQL);
                        if (this.m_ado.m_intError != 0)
                        {
                            break;
                        }
                    }
                }

                if (this.m_intError == 0 && this.m_ado.m_intError == 0)
                {
                    this.btnSave.Enabled = false;
                }
            }

        }

		private void btnDelete_Click(object sender, System.EventArgs e)
		{
			this.lstTreeDiam.SelectedItems[0].Remove();
			if (this.lstTreeDiam.Items.Count > 0)
			{
				this.lstTreeDiam.Items[this.lstTreeDiam.Items.Count-1].Selected=true;
			}
			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			this.lstTreeDiam.Focus();
		}

		private void lstTreeDiam_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.lstTreeDiam.SelectedItems.Count > 0)
			{
				if (this.lstTreeDiam.SelectedItems[0].Index ==
					this.lstTreeDiam.Items.Count-1)
				{
					if (this.btnDelete.Enabled==false) this.btnDelete.Enabled=true;
				}
				else
				{
					if (this.btnDelete.Enabled==true) this.btnDelete.Enabled=false;
				}
		
			}
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{

			this.ParentForm.Close();
			this.ParentForm.Dispose();
			
		}
	}
}
