using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_Rx.
	/// </summary>
	public class uc_rx_edit : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Panel panel1;
		public System.Windows.Forms.ListBox lstId;
		public System.Windows.Forms.ListBox lstMinor;
		public System.Windows.Forms.ListBox lstPackageMember;
		public System.Windows.Forms.TextBox txtDesc;
		public System.Windows.Forms.TextBox txtRxId;
		public System.Windows.Forms.ListBox lstMajor;
        public FIA_Biosum_Manager.ResizeFormUsingVisibleScrollBars m_oResizeForm = new ResizeFormUsingVisibleScrollBars();
		public System.Windows.Forms.TextBox txtCategory;
		public System.Windows.Forms.TextBox txtSubCategory;
		private System.Windows.Forms.Label lblSubCategory;
		private System.Windows.Forms.Label lblCategory;
		private System.Windows.Forms.Label lblId;
		private System.Windows.Forms.Label lblMinor;
		private System.Windows.Forms.Label lblMajor;
		private System.Windows.Forms.Button btnSelectRxId;
		private System.Windows.Forms.Label lblPackageMember;
		private System.Windows.Forms.Label lblDesc;
		private System.Windows.Forms.Label lblRxId;
		private FIA_Biosum_Manager.frmRxItem _frmRxItem=null;
		private ado_data_access m_oAdo = new ado_data_access();
		private Queries m_oQueries = new Queries();
		
		//private FIA_Biosum_Manager.RxItem_Collection _oRxItemCollection;
		
		
		

		//private ScrollBars _visibleScrollbars = ScrollBars.None;
		//public event EventHandler VisibleScrollbarsChanged;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_rx_edit()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			m_oResizeForm.ScrollBarParentControl=panel1;
			m_oResizeForm.MaximumWidth=770;
			m_oResizeForm.MaximumHeight=630;

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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtSubCategory = new System.Windows.Forms.TextBox();
            this.txtCategory = new System.Windows.Forms.TextBox();
            this.lblSubCategory = new System.Windows.Forms.Label();
            this.lblCategory = new System.Windows.Forms.Label();
            this.lblId = new System.Windows.Forms.Label();
            this.lblMinor = new System.Windows.Forms.Label();
            this.lblMajor = new System.Windows.Forms.Label();
            this.btnSelectRxId = new System.Windows.Forms.Button();
            this.lstId = new System.Windows.Forms.ListBox();
            this.lstMinor = new System.Windows.Forms.ListBox();
            this.lblPackageMember = new System.Windows.Forms.Label();
            this.lblDesc = new System.Windows.Forms.Label();
            this.lstPackageMember = new System.Windows.Forms.ListBox();
            this.txtDesc = new System.Windows.Forms.TextBox();
            this.txtRxId = new System.Windows.Forms.TextBox();
            this.lblRxId = new System.Windows.Forms.Label();
            this.lstMajor = new System.Windows.Forms.ListBox();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(755, 560);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.txtSubCategory);
            this.panel1.Controls.Add(this.txtCategory);
            this.panel1.Controls.Add(this.lblSubCategory);
            this.panel1.Controls.Add(this.lblCategory);
            this.panel1.Controls.Add(this.lblId);
            this.panel1.Controls.Add(this.lblMinor);
            this.panel1.Controls.Add(this.lblMajor);
            this.panel1.Controls.Add(this.btnSelectRxId);
            this.panel1.Controls.Add(this.lstId);
            this.panel1.Controls.Add(this.lstMinor);
            this.panel1.Controls.Add(this.lblPackageMember);
            this.panel1.Controls.Add(this.lblDesc);
            this.panel1.Controls.Add(this.lstPackageMember);
            this.panel1.Controls.Add(this.txtDesc);
            this.panel1.Controls.Add(this.txtRxId);
            this.panel1.Controls.Add(this.lblRxId);
            this.panel1.Controls.Add(this.lstMajor);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 16);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(749, 541);
            this.panel1.TabIndex = 0;
            this.panel1.Resize += new System.EventHandler(this.panel1_Resize);
            // 
            // txtSubCategory
            // 
            this.txtSubCategory.Location = new System.Drawing.Point(104, 240);
            this.txtSubCategory.Name = "txtSubCategory";
            this.txtSubCategory.Size = new System.Drawing.Size(328, 20);
            this.txtSubCategory.TabIndex = 33;
            // 
            // txtCategory
            // 
            this.txtCategory.Location = new System.Drawing.Point(104, 200);
            this.txtCategory.Name = "txtCategory";
            this.txtCategory.Size = new System.Drawing.Size(328, 20);
            this.txtCategory.TabIndex = 32;
            // 
            // lblSubCategory
            // 
            this.lblSubCategory.Location = new System.Drawing.Point(16, 240);
            this.lblSubCategory.Name = "lblSubCategory";
            this.lblSubCategory.Size = new System.Drawing.Size(80, 16);
            this.lblSubCategory.TabIndex = 31;
            this.lblSubCategory.Text = "Sub-Category:";
            // 
            // lblCategory
            // 
            this.lblCategory.Location = new System.Drawing.Point(16, 200);
            this.lblCategory.Name = "lblCategory";
            this.lblCategory.Size = new System.Drawing.Size(56, 24);
            this.lblCategory.TabIndex = 30;
            this.lblCategory.Text = "Category:";
            // 
            // lblId
            // 
            this.lblId.Location = new System.Drawing.Point(648, 8);
            this.lblId.Name = "lblId";
            this.lblId.Size = new System.Drawing.Size(80, 16);
            this.lblId.TabIndex = 29;
            this.lblId.Text = "Available Id\'s";
            // 
            // lblMinor
            // 
            this.lblMinor.Location = new System.Drawing.Point(240, 8);
            this.lblMinor.Name = "lblMinor";
            this.lblMinor.Size = new System.Drawing.Size(144, 16);
            this.lblMinor.TabIndex = 28;
            this.lblMinor.Text = "Sub-Category";
            // 
            // lblMajor
            // 
            this.lblMajor.Location = new System.Drawing.Point(10, 8);
            this.lblMajor.Name = "lblMajor";
            this.lblMajor.Size = new System.Drawing.Size(144, 16);
            this.lblMajor.TabIndex = 27;
            this.lblMajor.Text = "Category";
            // 
            // btnSelectRxId
            // 
            this.btnSelectRxId.Location = new System.Drawing.Point(648, 160);
            this.btnSelectRxId.Name = "btnSelectRxId";
            this.btnSelectRxId.Size = new System.Drawing.Size(88, 24);
            this.btnSelectRxId.TabIndex = 26;
            this.btnSelectRxId.Text = "Select";
            this.btnSelectRxId.Click += new System.EventHandler(this.button3_Click);
            // 
            // lstId
            // 
            this.lstId.Location = new System.Drawing.Point(648, 24);
            this.lstId.Name = "lstId";
            this.lstId.Size = new System.Drawing.Size(88, 134);
            this.lstId.TabIndex = 25;
            // 
            // lstMinor
            // 
            this.lstMinor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.lstMinor.Location = new System.Drawing.Point(232, 24);
            this.lstMinor.Name = "lstMinor";
            this.lstMinor.Size = new System.Drawing.Size(376, 134);
            this.lstMinor.TabIndex = 24;
            this.lstMinor.SelectedIndexChanged += new System.EventHandler(this.lstMinor_SelectedIndexChanged);
            // 
            // lblPackageMember
            // 
            this.lblPackageMember.Location = new System.Drawing.Point(493, 276);
            this.lblPackageMember.Name = "lblPackageMember";
            this.lblPackageMember.Size = new System.Drawing.Size(235, 16);
            this.lblPackageMember.TabIndex = 21;
            this.lblPackageMember.Text = "Packages that use this treatment";
            // 
            // lblDesc
            // 
            this.lblDesc.Location = new System.Drawing.Point(16, 276);
            this.lblDesc.Name = "lblDesc";
            this.lblDesc.Size = new System.Drawing.Size(96, 16);
            this.lblDesc.TabIndex = 20;
            this.lblDesc.Text = "Description";
            this.lblDesc.Click += new System.EventHandler(this.label2_Click);
            // 
            // lstPackageMember
            // 
            this.lstPackageMember.Location = new System.Drawing.Point(496, 296);
            this.lstPackageMember.Name = "lstPackageMember";
            this.lstPackageMember.Size = new System.Drawing.Size(240, 173);
            this.lstPackageMember.TabIndex = 19;
            // 
            // txtDesc
            // 
            this.txtDesc.Location = new System.Drawing.Point(16, 296);
            this.txtDesc.Multiline = true;
            this.txtDesc.Name = "txtDesc";
            this.txtDesc.Size = new System.Drawing.Size(448, 168);
            this.txtDesc.TabIndex = 18;
            // 
            // txtRxId
            // 
            this.txtRxId.Location = new System.Drawing.Point(104, 168);
            this.txtRxId.MaxLength = 3;
            this.txtRxId.Name = "txtRxId";
            this.txtRxId.Size = new System.Drawing.Size(88, 20);
            this.txtRxId.TabIndex = 17;
            this.txtRxId.TextChanged += new System.EventHandler(this.txtRxId_TextChanged);
            // 
            // lblRxId
            // 
            this.lblRxId.Location = new System.Drawing.Point(16, 168);
            this.lblRxId.Name = "lblRxId";
            this.lblRxId.Size = new System.Drawing.Size(80, 16);
            this.lblRxId.TabIndex = 16;
            this.lblRxId.Text = "Treatment ID:";
            this.lblRxId.Click += new System.EventHandler(this.label1_Click);
            // 
            // lstMajor
            // 
            this.lstMajor.Items.AddRange(new object[] {
            "Traditional Thinning  001-299",
            "Fuel Treatments 300-499",
            "Other Thinning Operations 650-699",
            "General Custom Defined 700-999"});
            this.lstMajor.Location = new System.Drawing.Point(12, 24);
            this.lstMajor.Name = "lstMajor";
            this.lstMajor.Size = new System.Drawing.Size(208, 134);
            this.lstMajor.TabIndex = 15;
            this.lstMajor.SelectedIndexChanged += new System.EventHandler(this.lstMajor_SelectedIndexChanged);
            // 
            // uc_rx_edit
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_rx_edit";
            this.Size = new System.Drawing.Size(755, 560);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		private void lstMajor_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.LoadSubCategoryListBox();


		}
		private void LoadSubCategoryListBox()
		{
			string str="";
			string strMajorItem="";

			this.lstMinor.Items.Clear();
			this.lstId.Items.Clear();

			strMajorItem = this.lstMajor.Text.Substring(0,this.lstMajor.Text.Length - 7);

			if (strMajorItem.Trim().ToUpper()=="GENERAL CUSTOM DEFINED")
			{
				LoadUnusedRxIdListBox(this.lstMajor.Text.Substring(this.lstMajor.Text.Length - 7,7));
			}
			else
			{

				this.m_oAdo.SqlQueryReader(this.m_oAdo.m_OleDbConnection,this.m_oQueries.m_oFvs.GetFVSSubCategoriesSQL(strMajorItem));

				if (m_oAdo.m_OleDbDataReader.HasRows)
				{
					while (m_oAdo.m_OleDbDataReader.Read())
					{
						if (m_oAdo.m_OleDbDataReader["desc"] != System.DBNull.Value)
						{
							str=Convert.ToString(m_oAdo.m_OleDbDataReader["desc"]).Trim() + " " + 
								Convert.ToString(m_oAdo.m_OleDbDataReader["min"]).Trim().PadLeft(3,'0') + "-" + 
								Convert.ToString(m_oAdo.m_OleDbDataReader["max"]).Trim().PadLeft(3,'0');
							this.lstMinor.Items.Add(str);
						}
					}
				}
				m_oAdo.m_OleDbDataReader.Close();
			}
		}
		private void LoadUnusedRxIdListBox(string p_strSelectedItem)
		{
			int x=0;
			int y=0;
			int intRxId=0;
			this.lstId.Items.Clear();
			string str = p_strSelectedItem;
			int intMin = Convert.ToInt32(str.Substring(0,3));
			int intMax = Convert.ToInt32(str.Substring(4,3));

			string[] strUsedRxIdArray = frmMain.g_oUtils.ConvertListToArray(this.ReferenceFormRxItem.UsedRxList,",");

			for (x=intMin;x<=intMax;x++)
			{
				if (this.ReferenceFormRxItem.UsedRxList.Trim().Length > 0)
				{
					for (y=0;y<=strUsedRxIdArray.Length-1;y++)
					{
						if (Convert.ToInt32(strUsedRxIdArray[y])==x)break;
						
					
					}
					if (y > strUsedRxIdArray.Length-1)
					{
						this.lstId.Items.Add(Convert.ToString(x).PadLeft(3,'0'));
					}
				}
				else
				{
					this.lstId.Items.Add(Convert.ToString(x).PadLeft(3,'0'));
				}
				
			}
		}

		private void lstMinor_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			LoadUnusedRxIdListBox(this.lstMinor.Text.Substring(this.lstMinor.Text.Length - 7,7));

		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			if (lstId.SelectedItems.Count==0) return;

			this.txtRxId.Text = this.lstId.SelectedItems[0].ToString().Trim();
			this.txtCategory.Text = this.lstMajor.SelectedItems[0].ToString().Trim().Substring(0,lstMajor.SelectedItems[0].ToString().Trim().Length-7);
			if (this.lstMinor.SelectedItems.Count > 0)
				this.txtSubCategory.Text = this.lstMinor.SelectedItems[0].ToString().Trim().Substring(0,lstMinor.SelectedItems[0].ToString().Trim().Length-7);
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			this.ParentForm.DialogResult = DialogResult.OK;
			this.ParentForm.Close();
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			this.ParentForm.DialogResult = DialogResult.Cancel;
			this.ParentForm.Close();
		}

		private void txtRxId_TextChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label2_Click(object sender, System.EventArgs e)
		{
		
		}

		private void panel1_Resize(object sender, System.EventArgs e)
		{
			
		}
		public void loadvalues()
		{
			
			if (this.ReferenceFormRxItem.m_strAction=="edit")
			{
				this.lblMajor.Hide(); this.lblMinor.Hide();this.lblId.Hide();
				this.lstMajor.Hide(); this.lstMinor.Hide();this.lstId.Hide();
				this.lblRxId.Top = this.lblMajor.Top;
				this.txtRxId.Top = this.lblRxId.Top;
				this.txtRxId.Left = 104;
				this.lblCategory.Top = this.lblRxId.Top + this.lblRxId.Height + 10;
				this.txtCategory.Top = this.lblCategory.Top;
				this.txtCategory.Left = 104;
				this.lblSubCategory.Top = this.lblCategory.Top + this.lblCategory.Height + 10;
				this.txtSubCategory.Top = this.lblSubCategory.Top;
				this.txtSubCategory.Left = 104;
				this.lblDesc.Top = this.lblSubCategory.Top + this.lblSubCategory.Height + 10;
				this.lblPackageMember.Top = this.lblDesc.Top;
				this.txtDesc.Top = this.lblDesc.Top + this.lblDesc.Height + 5;
				this.txtDesc.Height = this.ClientSize.Height - this.txtDesc.Top - 5;
				this.lstPackageMember.Top = this.txtDesc.Top;
				this.lstPackageMember.Height = this.txtDesc.Height;
				
				this.btnSelectRxId.Hide();
				this.lblId.Hide();
				this.txtRxId.Text = this.ReferenceFormRxItem.m_oRxItem.RxId;
				this.txtDesc.Text = this.ReferenceFormRxItem.m_oRxItem.Description;
				this.txtCategory.Text = this.ReferenceFormRxItem.m_oRxItem.Category;
				this.txtSubCategory.Text = this.ReferenceFormRxItem.m_oRxItem.SubCategory;

				this.txtRxId.Enabled=false;
			    this.txtCategory.Enabled=false;
				this.txtSubCategory.Enabled=false;

                if (this.ReferenceFormRxItem.m_oRxItem.RxPackageMemberList.Trim().Length > 0)
                {
                    string[] strArray = frmMain.g_oUtils.ConvertListToArray(ReferenceFormRxItem.m_oRxItem.RxPackageMemberList, ",");
                    for (int x = 0; x <= strArray.Length - 1; x++)
                    {
                        this.lstPackageMember.Items.Add(strArray[x]);
                    }
                }





			}
			else
			{
				this.m_oQueries.m_oFvs.LoadDatasource=true;
				this.m_oQueries.LoadDatasources(true);
				this.m_oAdo.OpenConnection(m_oAdo.getMDBConnString(this.m_oQueries.m_strTempDbFile,"",""));
				//
				//populate category list box
				//
				this.LoadCategoryListBox();


			}
		}
		public void savevalues()
		{
			if (this.txtRxId.Text.Trim().Length == 0)
			{
				this.ReferenceFormRxItem.m_strError="Select a treatment";
				this.ReferenceFormRxItem.m_intError=-1;
				return;
			}
			if (this.ReferenceFormRxItem.m_strAction=="new")
			{
				this.ReferenceFormRxItem.m_oRxItem.RxId=this.txtRxId.Text;
				this.ReferenceFormRxItem.m_oRxItem.Category = this.txtCategory.Text;
				this.ReferenceFormRxItem.m_oRxItem.SubCategory=this.txtSubCategory.Text;
				this.ReferenceFormRxItem.m_oRxItem.Description=this.txtDesc.Text;
			}
			else
			{
				this.ReferenceFormRxItem.m_oRxItem.Description=this.txtDesc.Text;
			}
		    
		}
		private void LoadCategoryListBox()
		{
			string str="";
			this.lstMajor.Items.Clear();
			this.m_oAdo.m_strSQL = "SELECT * FROM " + this.m_oQueries.m_oFvs.m_strFvsCatTable;
			
			m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection,m_oAdo.m_strSQL);

			if (m_oAdo.m_OleDbDataReader.HasRows)
			{
				while (m_oAdo.m_OleDbDataReader.Read())
				{
					if (m_oAdo.m_OleDbDataReader["desc"] != System.DBNull.Value)
					{
						str=Convert.ToString(m_oAdo.m_OleDbDataReader["desc"]) + " " + 
							Convert.ToString(m_oAdo.m_OleDbDataReader["min"]).Trim().PadLeft(3,'0') + "-" + 
							Convert.ToString(m_oAdo.m_OleDbDataReader["max"]).Trim().PadLeft(3,'0');
						this.lstMajor.Items.Add(str);
					}
				}
			}
			m_oAdo.m_OleDbDataReader.Close();

		

		}

		public FIA_Biosum_Manager.frmRxItem ReferenceFormRxItem
		{
			get {return this._frmRxItem;}
			set {this._frmRxItem=value;}

		}
		

		
	}
}
