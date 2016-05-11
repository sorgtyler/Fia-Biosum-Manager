using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_rx_package_fvscmd_list.
	/// </summary>
	public class uc_rx_package_fvscmd_list : System.Windows.Forms.UserControl
	{
		public int m_intError=0;
		public string m_strError="";
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.ListView lvRxPackageFVSCmd;
		private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvAlternateColors = new ListViewAlternateBackgroundColors();
		private System.Windows.Forms.Button btnUp;
		private System.Windows.Forms.Button btnDown;
		
		private FIA_Biosum_Manager.frmRxPackageItem _frmRxPackageItem=null;
		
		private int m_intCurrSelect=-1;
		//rx package edit
		const int COLUMN_NULL=0;
		const int COLUMN_PACKAGE=1;
		const int COLUMN_RX=2;
		const int COLUMN_FVSCMD=3;
		const int COLUMN_P1=4;
		const int COLUMN_P2=5;
		const int COLUMN_P3=6;
		const int COLUMN_P4=7;
		const int COLUMN_P5=8;
		const int COLUMN_P6=9;
		const int COLUMN_P7=10
		const int COLUMN_OTHER=11;
		
		


		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_rx_package_fvscmd_list()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			//m_oResizeForm.ScrollBarParentControl=panel1;

			// TODO: Add any initialization after the InitializeComponent call

			m_oLvAlternateColors.ReferenceAlternateBackgroundColor=frmMain.g_oGridViewAlternateRowBackgroundColor;
			this.m_oLvAlternateColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
			this.m_oLvAlternateColors.ReferenceListView=this.lvRxPackageFVSCmd;
			this.m_oLvAlternateColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
			this.m_oLvAlternateColors.CustomFullRowSelect=true;
			if (frmMain.g_oGridViewFont != null) this.lvRxPackageFVSCmd.Font = frmMain.g_oGridViewFont;


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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(uc_rx_package_fvscmd_list));
			this.panel1 = new System.Windows.Forms.Panel();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnDown = new System.Windows.Forms.Button();
			this.btnUp = new System.Windows.Forms.Button();
			this.lvRxPackageFVSCmd = new System.Windows.Forms.ListView();
			this.panel1.SuspendLayout();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.groupBox1);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel1.Location = new System.Drawing.Point(0, 0);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(688, 424);
			this.panel1.TabIndex = 0;
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btnDown);
			this.groupBox1.Controls.Add(this.btnUp);
			this.groupBox1.Controls.Add(this.lvRxPackageFVSCmd);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(688, 424);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// btnDown
			// 
			this.btnDown.Image = ((System.Drawing.Image)(resources.GetObject("btnDown.Image")));
			this.btnDown.Location = new System.Drawing.Point(20, 396);
			this.btnDown.Name = "btnDown";
			this.btnDown.Size = new System.Drawing.Size(48, 24);
			this.btnDown.TabIndex = 25;
			this.btnDown.Click += new System.EventHandler(this.btnDown_Click);
			// 
			// btnUp
			// 
			this.btnUp.Image = ((System.Drawing.Image)(resources.GetObject("btnUp.Image")));
			this.btnUp.Location = new System.Drawing.Point(20, 372);
			this.btnUp.Name = "btnUp";
			this.btnUp.Size = new System.Drawing.Size(48, 24);
			this.btnUp.TabIndex = 24;
			this.btnUp.Click += new System.EventHandler(this.btnUp_Click);
			// 
			// lvRxPackageFVSCmd
			// 
			this.lvRxPackageFVSCmd.GridLines = true;
			this.lvRxPackageFVSCmd.Location = new System.Drawing.Point(16, 24);
			this.lvRxPackageFVSCmd.MultiSelect = false;
			this.lvRxPackageFVSCmd.Name = "lvRxPackageFVSCmd";
			this.lvRxPackageFVSCmd.Size = new System.Drawing.Size(656, 344);
			this.lvRxPackageFVSCmd.TabIndex = 17;
			this.lvRxPackageFVSCmd.View = System.Windows.Forms.View.Details;
			this.lvRxPackageFVSCmd.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lvRxPackageFVSCmd_MouseUp);
			this.lvRxPackageFVSCmd.SelectedIndexChanged += new System.EventHandler(this.lvRxPackageFVSCmd_SelectedIndexChanged);
			// 
			// uc_rx_package_fvscmd_list
			// 
			this.Controls.Add(this.panel1);
			this.Name = "uc_rx_package_fvscmd_list";
			this.Size = new System.Drawing.Size(688, 424);
			this.Resize += new System.EventHandler(this.uc_rx_package_fvscmd_list_Resize);
			this.panel1.ResumeLayout(false);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void uc_rx_package_fvscmd_list_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.lvRxPackageFVSCmd.Left = 5;
				this.lvRxPackageFVSCmd.Width = this.Width - 10;

				
				btnDown.Top = this.ClientSize.Height - btnDown.Height - 5;
				btnUp.Top = btnDown.Top - btnDown.Height;
				


				this.lvRxPackageFVSCmd.Height = this.btnUp.Top - this.lvRxPackageFVSCmd.Top - 10;
	
			}
			catch
			{
			}
		}
		private void btnUp_Click(object sender, EventArgs e)
		{
			if (this.lvRxPackageFVSCmd.SelectedItems.Count == 0) return;
			if (this.lvRxPackageFVSCmd.SelectedItems[0].Index == 0) return;

			int intSaveCurrSelect = this.m_intCurrSelect;
			
			int int1=0;
			int int2=0;
			int x,y;

 



			string[] strValuesArray = new string[this.lvRxPackageFVSCmd.Columns.Count];
			//load into an array the values from the selected row
			for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
			{
				strValuesArray[x] = this.lvRxPackageFVSCmd.SelectedItems[0].SubItems[x].Text;
			}
			//swap values between the selected row and the row above
			y = this.lvRxPackageFVSCmd.SelectedItems[0].Index;
			for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
			{
				this.lvRxPackageFVSCmd.SelectedItems[0].SubItems[x].Text =
					this.lvRxPackageFVSCmd.Items[y - 1].SubItems[x].Text;

				this.lvRxPackageFVSCmd.Items[y - 1].SubItems[x].Text = 
					strValuesArray[x];
			}
			//adjust the selected row
			this.lvRxPackageFVSCmd.Items[y - 1].Selected = true;
			this.lvRxPackageFVSCmd.Focus();

		}

		private void btnDown_Click(object sender, EventArgs e)
		{
			if (this.lvRxPackageFVSCmd.SelectedItems.Count == 0) return;
			if (this.lvRxPackageFVSCmd.SelectedItems[0].Index == this.lvRxPackageFVSCmd.Items.Count-1) return;

			int intSaveCurrSelect = this.m_intCurrSelect;
			string str1;
			string str2;
			int int1=0;
			int int2=0;
			int x, y;



			string[] strValuesArray = new string[this.lvRxPackageFVSCmd.Columns.Count];
			//load into an array the values from the selected row
			for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
			{
				strValuesArray[x] = this.lvRxPackageFVSCmd.SelectedItems[0].SubItems[x].Text;
			}
			//swap values between the selected row and the row below
			y = this.lvRxPackageFVSCmd.SelectedItems[0].Index;
			for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
			{
				this.lvRxPackageFVSCmd.SelectedItems[0].SubItems[x].Text =
					this.lvRxPackageFVSCmd.Items[y + 1].SubItems[x].Text;

				this.lvRxPackageFVSCmd.Items[y + 1].SubItems[x].Text =
					strValuesArray[x];
			}
			//adjust the selected row
			this.lvRxPackageFVSCmd.Items[y + 1].Selected = true;
			this.lvRxPackageFVSCmd.Focus();
		}
		public void loadvalues()
		{
			int intRowCount=0;
			int x,y;
			
			this.lvRxPackageFVSCmd.Clear();
			this.m_oLvAlternateColors.InitializeRowCollection();      
			
			
			this.lvRxPackageFVSCmd.Columns.Add(" ",2,HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("Package", 60, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("Rx", 60, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("FVSCmd", 100, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("P1", 50, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("P2", 50, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("P3", 50, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("P4", 50, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("P5", 50, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("P6", 50, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("P7", 50, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("Other", 100, HorizontalAlignment.Left);

			
			
			this.m_oLvAlternateColors.ListView();

			if (this.lvRxPackageFVSCmd.Items.Count > 0) this.lvRxPackageFVSCmd.Items[0].Selected=true;
			
		}
		public void savevalues()
		{
			
			
		}

		private void lvRxPackageFVSCmd_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.ReferenceFormRxItem.SetToolBarFVSCommandButtonsEnabled("Y","Y","Y","Y","Y");
			
			if (this.lvRxPackageFVSCmd.SelectedItems.Count > 0)
			{
				m_intCurrSelect = lvRxPackageFVSCmd.SelectedItems[0].Index;
				this.m_oLvAlternateColors.DelegateListViewItem(lvRxPackageFVSCmd.SelectedItems[0]);
			}
		}

		private void lvRxPackageFVSCmd_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			try
			{
				if (e.Button == MouseButtons.Left)
				{
					int intRowHt = this.lvRxPackageFVSCmd.Items[0].Bounds.Height;
					double dblRow = (double)(e.Y / intRowHt);
					this.lvRxPackageFVSCmd.Items[this.lvRxPackageFVSCmd.TopItem.Index + (int)dblRow-1].Selected=true;
				}
			}
			catch 
			{
			}
		}
		public void EditItem()
		{
			int x;
			if (this.lvRxPackageFVSCmd.SelectedItems.Count == 0) return;
		}
		private void UpdateListViewRxItem(FIA_Biosum_Manager.RxItemFvsCommandItem p_oRxItemFvsCommandItem)
		{
			if (lvRxPackageFVSCmd.SelectedItems.Count==0) return;
			int x=lvRxPackageFVSCmd.SelectedItems[0].Index;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_RX].Text=p_oRxItemFvsCommandItem.RxId;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_FVSCMD].Text=p_oRxItemFvsCommandItem.FVSCommand;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_P1].Text=p_oRxItemFvsCommandItem.Parameter1;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_P2].Text=p_oRxItemFvsCommandItem.Parameter2;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_P3].Text=p_oRxItemFvsCommandItem.Parameter3;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_P4].Text=p_oRxItemFvsCommandItem.Parameter4;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_P5].Text=p_oRxItemFvsCommandItem.Parameter5;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_P6].Text=p_oRxItemFvsCommandItem.Parameter6;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_P7].Text=p_oRxItemFvsCommandItem.Parameter7;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_OTHER].Text=p_oRxItemFvsCommandItem.Other;
		}
		private void UpdateListViewRxItem(FIA_Biosum_Manager.RxItemFvsCommandItem p_oRxItemFvsCommandItem,int p_intRow)
		{
			this.lvRxPackageFVSCmd.BeginUpdate();
			
				this.lvRxPackageFVSCmd.Items[p_intRow].SubItems[COLUMN_RX].Text=p_oRxItemFvsCommandItem.RxId;
				this.lvRxPackageFVSCmd.Items[p_intRow].SubItems[COLUMN_FVSCMD].Text=p_oRxItemFvsCommandItem.FVSCommand;
				this.lvRxPackageFVSCmd.Items[p_intRow].SubItems[COLUMN_P1].Text=p_oRxItemFvsCommandItem.Parameter1;
				this.lvRxPackageFVSCmd.Items[p_intRow].SubItems[COLUMN_P2].Text=p_oRxItemFvsCommandItem.Parameter2;
				this.lvRxPackageFVSCmd.Items[p_intRow].SubItems[COLUMN_P3].Text=p_oRxItemFvsCommandItem.Parameter3;
				this.lvRxPackageFVSCmd.Items[p_intRow].SubItems[COLUMN_P4].Text=p_oRxItemFvsCommandItem.Parameter4;
				this.lvRxPackageFVSCmd.Items[p_intRow].SubItems[COLUMN_P5].Text=p_oRxItemFvsCommandItem.Parameter5;
				this.lvRxPackageFVSCmd.Items[p_intRow].SubItems[COLUMN_P6].Text=p_oRxItemFvsCommandItem.Parameter6;
				this.lvRxPackageFVSCmd.Items[p_intRow].SubItems[COLUMN_P7].Text=p_oRxItemFvsCommandItem.Parameter7;
				this.lvRxPackageFVSCmd.Items[p_intRow].SubItems[COLUMN_OTHER].Text=p_oRxItemFvsCommandItem.Other;
			
			this.lvRxPackageFVSCmd.EndUpdate();

		}
		public void RemoveItem()
		{
			if (this.lvRxPackageFVSCmd.SelectedItems.Count == 0) return;
			int x;
			int y;
			int intSelect;
			/**********************************************
			 **lets see if we have one to remove
			 **********************************************/
			int index=this.m_intCurrSelect;
			intSelect=index;

			
								
		}
		
		
		public void AddItem()
		{
		}
		private void btnEdit_Click(object sender, System.EventArgs e)
		{
			EditItem();
		}
		private void AddItemToList(FIA_Biosum_Manager.RxItemFvsCommandItem oItem)
		{
			
			this.lvRxPackageFVSCmd.Items.Add("");
			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].UseItemStyleForSubItems=false;
			for (int z=1;z<=this.lvRxPackageFVSCmd.Columns.Count-1;z++)
			{
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems.Add(" ");
			}
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_RX].Text=oItem.RxId;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_FVSCMD].Text=oItem.FVSCommand;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P1].Text=oItem.Parameter1;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P2].Text=oItem.Parameter2;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P3].Text=oItem.Parameter3;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P4].Text=oItem.Parameter4;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P5].Text=oItem.Parameter5;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P6].Text=oItem.Parameter6;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P7].Text=oItem.Parameter7;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_OTHER].Text=oItem.Other;

									
			this.m_oLvAlternateColors.AddRow();
			this.m_oLvAlternateColors.AddColumns(lvRxPackageFVSCmd.Items.Count-1,this.lvRxPackageFVSCmd.Columns.Count);

			this.ReferenceFormRxItem.SetToolBarFVSCommandButtonsEnabled("","Y","Y","Y","");

			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].Selected=true;
		}
		
		public void RemoveAllItems()
		{
			if (lvRxPackageFVSCmd.SelectedItems.Count==0) return;

			this.lvRxPackageFVSCmd.Items.Clear();
			this.m_oLvAlternateColors.InitializeRowCollection();

			for (int x=ReferenceFormRxItem.m_oRxItem.ReferenceFvsCommandsCollection.Count-1;x>=0;x--)
			     ReferenceFormRxItem.m_oRxItem.ReferenceFvsCommandsCollection.Remove(x);

			this.ReferenceFormRxItem.SetToolBarFVSCommandButtonsEnabled("Y","N","N","N","Y");

		}
			
		public FIA_Biosum_Manager.frmRxPackageItem ReferenceFormRxPackageItem
		{
			get {return this._frmRxPackageItem;}
			set {this._frmRxPackageItem=value;}

		}

	}

}
