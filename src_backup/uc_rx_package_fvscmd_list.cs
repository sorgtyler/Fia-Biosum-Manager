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
		const int COLUMN_CYCLE=3;
		const int COLUMN_FVSCMD=4;
		const int COLUMN_P1=5;
		const int COLUMN_P2=6;
		const int COLUMN_P3=7;
		const int COLUMN_P4=8;
		const int COLUMN_P5=9;
		const int COLUMN_P6=10;
		const int COLUMN_P7=11;
		const int COLUMN_OTHER=12;
		const int COLUMN_FVSCMDID=13;
        private Label lblDesc;

        private ListViewColumnSorter lvwColumnSorter;
		
		


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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_rx_package_fvscmd_list));
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnDown = new System.Windows.Forms.Button();
            this.btnUp = new System.Windows.Forms.Button();
            this.lvRxPackageFVSCmd = new System.Windows.Forms.ListView();
            this.lblDesc = new System.Windows.Forms.Label();
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
            this.groupBox1.Controls.Add(this.lblDesc);
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
            this.lvRxPackageFVSCmd.Location = new System.Drawing.Point(16, 45);
            this.lvRxPackageFVSCmd.MultiSelect = false;
            this.lvRxPackageFVSCmd.Name = "lvRxPackageFVSCmd";
            this.lvRxPackageFVSCmd.Size = new System.Drawing.Size(656, 323);
            this.lvRxPackageFVSCmd.TabIndex = 17;
            this.lvRxPackageFVSCmd.UseCompatibleStateImageBehavior = false;
            this.lvRxPackageFVSCmd.View = System.Windows.Forms.View.Details;
            this.lvRxPackageFVSCmd.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.lvRxPackageFVSCmd_ColumnClick);
            this.lvRxPackageFVSCmd.SelectedIndexChanged += new System.EventHandler(this.lvRxPackageFVSCmd_SelectedIndexChanged);
            this.lvRxPackageFVSCmd.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lvRxPackageFVSCmd_MouseUp);
            // 
            // lblDesc
            // 
            this.lblDesc.AutoSize = true;
            this.lblDesc.Location = new System.Drawing.Point(6, 16);
            this.lblDesc.Name = "lblDesc";
            this.lblDesc.Size = new System.Drawing.Size(620, 26);
            this.lblDesc.TabIndex = 26;
            this.lblDesc.Text = "This documentation screen provides a place to view and manually enter treatment f" +
    "vs commands represented in the KCP/KEY file.\r\n.\r\n";
            // 
            // uc_rx_package_fvscmd_list
            // 
            this.Controls.Add(this.panel1);
            this.Name = "uc_rx_package_fvscmd_list";
            this.Size = new System.Drawing.Size(688, 424);
            this.Resize += new System.EventHandler(this.uc_rx_package_fvscmd_list_Resize);
            this.panel1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
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
			MoveUp();
			
		}
		private void MoveUp()
		{
			if (this.lvRxPackageFVSCmd.SelectedItems.Count == 0) return;
			if (this.lvRxPackageFVSCmd.SelectedItems[0].Index == 0) return;

			int intSaveCurrSelect = this.m_intCurrSelect;
			
			int int1=-1;
			int int2=-1;
			int x,y;
			int intCount=0;
			int intReplaceIndex=0;
			string strRx="";
			string strCycle="";
			for (x=m_intCurrSelect-1;x>=0;x--)
			{
				if (this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_RX].Text.Trim().Length > 0)
				{
					if (strRx.Trim().Length > 0 && strCycle.Trim().Length > 0 && 
						(lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_RX].Text.Trim() != strRx ||
						lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text.Trim() != strCycle)) break;
					else intCount++;
					strRx = lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_RX].Text.Trim();
					strCycle=lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text.Trim();
				}
				else 
				{
					if (strRx.Trim().Length > 0) break;
					intCount++;
					break;
				}
				

			}
			intReplaceIndex = m_intCurrSelect - intCount;
			//intReplaceIndex=ReplaceIndex;


 
					
				if (this.ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection!=null)
				{
					for (x=0;x<=ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Count-1;x++)
					{
						if (ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(x).Index==intSaveCurrSelect)
						{
							int2=x;
						}
						else
						{
							if (ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(x).Index==intSaveCurrSelect-1)
							{
								int1=x;
							}
						}
					}

					//swap locations with the one directly above the currently selected item
					if (int1 > -1)
						ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(int1).Index = intSaveCurrSelect;
					if (int2 > -1)
						ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(int2).Index = intSaveCurrSelect-1;
				
				}
				



			string[] strValuesArray = new string[this.lvRxPackageFVSCmd.Columns.Count];
			
			
			//load into an array the values from the selected row
			for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
			{
				strValuesArray[x] = this.lvRxPackageFVSCmd.SelectedItems[0].SubItems[x].Text;
			}


			//swap values between the selected row and the row above
			y = this.lvRxPackageFVSCmd.SelectedItems[0].Index;
			for (int z=this.lvRxPackageFVSCmd.SelectedItems[0].Index;z>intReplaceIndex;z--)
			{
				for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
				{
					this.lvRxPackageFVSCmd.Items[z].SubItems[x].Text =
						this.lvRxPackageFVSCmd.Items[z - 1].SubItems[x].Text;

					
				}
			}
			int index = m_intCurrSelect - intReplaceIndex;
			for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
			{

				this.lvRxPackageFVSCmd.Items[intReplaceIndex].SubItems[x].Text = 
					strValuesArray[x];
			}
			//adjust the selected row
			this.lvRxPackageFVSCmd.Items[intReplaceIndex].Selected = true;


			for (x=0;x<=ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Count-1;x++)
			{
				for (y=0;y<=this.lvRxPackageFVSCmd.Items.Count-1;y++)
				{
					if (this.lvRxPackageFVSCmd.Items[y].SubItems[COLUMN_PACKAGE].Text.Trim() ==
						ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(x).RxPackageId.Trim() &&
						this.lvRxPackageFVSCmd.Items[y].SubItems[COLUMN_FVSCMD].Text.Trim() == 
						ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(x).FVSCommand.Trim() &&
						this.lvRxPackageFVSCmd.Items[y].SubItems[COLUMN_FVSCMDID].Text.Trim() == 
						Convert.ToString(ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(x).FVSCommandId).Trim())
					{
						this.ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(x).ListViewIndex=y;
					}
				}
			}

			this.lvRxPackageFVSCmd.Focus();

		}

		private void btnDown_Click(object sender, EventArgs e)
		{
			MoveDown();
		}
		private void MoveDown()
		{
			if (this.lvRxPackageFVSCmd.SelectedItems.Count == 0) return;
			if (this.lvRxPackageFVSCmd.SelectedItems[0].Index == this.lvRxPackageFVSCmd.Items.Count-1) return;

			int intSaveCurrSelect = this.m_intCurrSelect;
			string str1;
			string str2;
			int int1=-1;
			int int2=-1;
			int x, y;

			int intCount=0;
			int intReplaceIndex=0;
			string strRx="";
			string strCycle="";
			for (x=m_intCurrSelect+1;x<=lvRxPackageFVSCmd.Items.Count-1;x++)
			{
				if (this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_RX].Text.Trim().Length > 0)
				{
					if (strRx.Trim().Length > 0 && strCycle.Trim().Length > 0 && 
						(lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_RX].Text.Trim() != strRx ||
						lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text.Trim() != strCycle)) break;
					else intCount++;
					strRx = lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_RX].Text.Trim();
					strCycle=lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text.Trim();
				}
				else 
				{
					if (strRx.Trim().Length > 0) break;
					intCount++;
					break;
				}
				

			}
			intReplaceIndex = m_intCurrSelect + intCount;

			
			if (ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection!=null)
			{
				for (x=0;x<=ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Count-1;x++)
				{
					if (ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(x).Index==intSaveCurrSelect)
					{
						int2=x;
					}
					else
					{
						if (ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(x).Index==intSaveCurrSelect+1)
						{
							int1=x;
						}
					}
				}

				//swap locations with the one directly above the currently selected item
				if (int1 > -1)
					ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(int1).Index = intSaveCurrSelect;
				if (int2 > -1)
					ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(int2).Index = intSaveCurrSelect+1;
			}
			


			string[] strValuesArray = new string[this.lvRxPackageFVSCmd.Columns.Count];
			//load into an array the values from the selected row
			for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
			{
				strValuesArray[x] = this.lvRxPackageFVSCmd.SelectedItems[0].SubItems[x].Text;
			}

			//swap values between the selected row and the row below
			y = this.lvRxPackageFVSCmd.SelectedItems[0].Index;
			for (int z=this.lvRxPackageFVSCmd.SelectedItems[0].Index;z<intReplaceIndex;z++)
			{
				for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
				{
					this.lvRxPackageFVSCmd.Items[z].SubItems[x].Text =
						this.lvRxPackageFVSCmd.Items[z + 1].SubItems[x].Text;

					
				}
			}
			
			for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
			{

				this.lvRxPackageFVSCmd.Items[intReplaceIndex].SubItems[x].Text = 
					strValuesArray[x];
			}
			//adjust the selected row
			this.lvRxPackageFVSCmd.Items[intReplaceIndex].Selected = true;

			for (x=0;x<=ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Count-1;x++)
			{
				for (y=0;y<=this.lvRxPackageFVSCmd.Items.Count-1;y++)
				{
					if (this.lvRxPackageFVSCmd.Items[y].SubItems[COLUMN_PACKAGE].Text.Trim() ==
						ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(x).RxPackageId.Trim() &&
						this.lvRxPackageFVSCmd.Items[y].SubItems[COLUMN_FVSCMD].Text.Trim() == 
						ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(x).FVSCommand.Trim() &&
						this.lvRxPackageFVSCmd.Items[y].SubItems[COLUMN_FVSCMDID].Text.Trim() == 
						Convert.ToString(ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(x).FVSCommandId).Trim())
					{
						this.ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(x).Index=y;
					}
				}
			}
			this.lvRxPackageFVSCmd.Focus();
			//swap values between the selected row and the row below
			//y = this.lvRxPackageFVSCmd.SelectedItems[0].Index;
			//for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
			//{
			//	this.lvRxPackageFVSCmd.SelectedItems[0].SubItems[x].Text =
			//		this.lvRxPackageFVSCmd.Items[y + 1].SubItems[x].Text;
//
//				this.lvRxPackageFVSCmd.Items[y + 1].SubItems[x].Text =
//					strValuesArray[x];
//			}
			//adjust the selected row
//			this.lvRxPackageFVSCmd.Items[y + 1].Selected = true;
//			this.lvRxPackageFVSCmd.Focus();

		}
		private void MoveDown(int p_intItemToMoveIndex, int p_intPositionToMoveIndex)
		{
			    if (this.lvRxPackageFVSCmd.Items.Count==0) return;
				if (this.lvRxPackageFVSCmd.Items.Count-1==p_intItemToMoveIndex) return;
				
				
				string str1;
				string str2;
				int int1=-1;
				int int2=-1;
				int x, y;

				int intCount=0;
				int intReplaceIndex=0;
				string strRx="";
				string strCycle="";
						
			
				string[] strValuesArray = new string[this.lvRxPackageFVSCmd.Columns.Count];
				//load into an array the values from the  row to move
				for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
				{
					strValuesArray[x] = this.lvRxPackageFVSCmd.Items[p_intItemToMoveIndex].SubItems[x].Text;
				}

				//swap values between the selected row and the row below
				for (int z=lvRxPackageFVSCmd.Items[p_intItemToMoveIndex].Index;z<p_intPositionToMoveIndex;z++)
				{
					for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
					{
						this.lvRxPackageFVSCmd.Items[z].SubItems[x].Text =
							this.lvRxPackageFVSCmd.Items[z + 1].SubItems[x].Text;

					
					}
				}
			
				for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
				{

					this.lvRxPackageFVSCmd.Items[p_intPositionToMoveIndex].SubItems[x].Text = 
						strValuesArray[x];
				}
				
			

		}
		private void MoveUp(int p_intItemToMoveIndex, int p_intPositionToMoveIndex)
		{
			if (this.lvRxPackageFVSCmd.Items.Count==0) return;
			if (p_intItemToMoveIndex==0) return;
				
				
			string str1;
			string str2;
			int int1=-1;
			int int2=-1;
			int x, y;

			int intCount=0;
			int intReplaceIndex=0;
			string strRx="";
			string strCycle="";
						
			
			string[] strValuesArray = new string[this.lvRxPackageFVSCmd.Columns.Count];
			//load into an array the values from the  row to move
			for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
			{
				strValuesArray[x] = this.lvRxPackageFVSCmd.Items[p_intItemToMoveIndex].SubItems[x].Text;
			}

			//swap values between the selected row and the row below
			for (int z=lvRxPackageFVSCmd.Items[p_intItemToMoveIndex].Index;z>p_intPositionToMoveIndex;z--)
			{
				for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
				{
					this.lvRxPackageFVSCmd.Items[z].SubItems[x].Text =
						this.lvRxPackageFVSCmd.Items[z - 1].SubItems[x].Text;

					
				}
			}
			
			for (x = 0; x <= this.lvRxPackageFVSCmd.Columns.Count - 1; x++)
			{

				this.lvRxPackageFVSCmd.Items[p_intPositionToMoveIndex].SubItems[x].Text = 
					strValuesArray[x];
			}
				
			

		}
		
		public void loadvalues()
		{
			int x,y,z;
			this.lvRxPackageFVSCmd.Clear();
			this.m_oLvAlternateColors.InitializeRowCollection();      
			
			
			this.lvRxPackageFVSCmd.Columns.Add(" ",2,HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("Package", 60, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("Rx", 60, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("SimCycle", 100, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("FVSCmd", 100, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("P1", 50, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("P2", 50, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("P3", 50, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("P4", 50, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("P5", 50, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("P6", 50, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("P7", 50, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("Other", 100, HorizontalAlignment.Left);
			this.lvRxPackageFVSCmd.Columns.Add("FVSCmdId",50, HorizontalAlignment.Left);

            lvwColumnSorter = new ListViewColumnSorter();
            this.lvRxPackageFVSCmd.ListViewItemSorter = lvwColumnSorter;

			if (this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1 != null &&
				this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1.Count > 0)
			{
				for (x=0;x<=this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1.Count-1;x++)
				{
					if (this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1.Item(x).RxId.Trim().Length == 0)
					{
						for (y=0;y<=this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oFvsCommandItem_Collection1.Count-1;y++)
						{
							if (this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oFvsCommandItem_Collection1.Item(y).FVSCommand.Trim().ToUpper()==
								this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1.Item(x).FVSCommand.Trim().ToUpper() &&
								this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oFvsCommandItem_Collection1.Item(y).FVSCommandId ==
								this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1.Item(x).FVSCommandId)
							{
								this.AddItemToList(ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(y));
							}
						}
					}
					else
					{
						if (this.ReferenceFormRxPackageItem.ReferenceRxItemCollection != null)
						{
							for (y=0;y<=this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Count - 1;y++)
							{
								if (ReferenceFormRxPackageItem.m_oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1.Item(x).RxId.Trim()==
									ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(y).RxId.Trim())

								{
									for (z=0;z<=ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(y).ReferenceFvsCommandsCollection.Count-1;z++)
									{
										if (ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(y).ReferenceFvsCommandsCollection.Item(z).FVSCommand.Trim().ToUpper() ==
											ReferenceFormRxPackageItem.m_oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1.Item(x).FVSCommand.Trim().ToUpper() &&
											ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(y).ReferenceFvsCommandsCollection.Item(z).FVSCommandId ==
											ReferenceFormRxPackageItem.m_oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1.Item(x).FVSCommandId)
										{
											this.AddItemToList(
												ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(y).ReferenceFvsCommandsCollection.Item(z),
												ReferenceFormRxPackageItem.m_oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1.Item(x).FVSCycle);
											break;
										}

									}
									break;

								}
							}
						}
					}
				}
					
						
			}
			//
			//double check to make sure all the rx fvs items are added
			//
			
				//
				//1st simulation
				//
				for (x=0;x<=this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Count - 1;x++)
				{
					ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).FvsCycleList="";

					if (ReferenceFormRxPackageItem.ReferenceRxPackageItem.SimulationYear1Rx.Trim()==
						ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim())
					{
						if (ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection != null &&
							ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Count > 0)
						{
							for (y=0;y<=ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Count-1;y++)
							{
								//
								//make sure the rx,fvscmd and fvscmdid are not added already
								//
								for (z=0;z<=this.lvRxPackageFVSCmd.Items.Count-1;z++)
								{
									if (this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_RX].Text.Trim().Length > 0)
									{
										if (this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_RX].Text.Trim() ==
											ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim() &&
											this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_FVSCMD].Text.Trim().ToUpper() ==
											ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Item(y).FVSCommand.Trim().ToUpper() &&
											this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_FVSCMDID].Text.Trim() == 
											Convert.ToString(ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Item(y).FVSCommandId).Trim() &&
											this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_CYCLE].Text.Trim() == 
											ReferenceFormRxPackageItem.ReferenceRxPackageItem.SimulationYear1Fvs.Trim())
											   break;




									}
								}
								if (z>this.lvRxPackageFVSCmd.Items.Count-1)
								{
									this.AddItemToList(
										ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Item(y),
										ReferenceFormRxPackageItem.ReferenceRxPackageItem.SimulationYear1Fvs);
								}
							}
						}
						break;
					}
				}
				//
				//2nd simulation
				//
				for (x=0;x<=this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Count - 1;x++)
				{
					ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).FvsCycleList="";

					if (ReferenceFormRxPackageItem.ReferenceRxPackageItem.SimulationYear2Rx.Trim()==
						ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim())
					{
						if (ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection != null &&
							ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Count > 0)
						{
							for (y=0;y<=ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Count-1;y++)
							{
								//
								//make sure the rx,fvscmd and fvscmdid are not added already
								//
								for (z=0;z<=this.lvRxPackageFVSCmd.Items.Count-1;z++)
								{
									if (this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_RX].Text.Trim().Length > 0)
									{
										if (this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_RX].Text.Trim() ==
											ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim() &&
											this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_FVSCMD].Text.Trim().ToUpper() ==
											ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Item(y).FVSCommand.Trim().ToUpper() &&
											this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_FVSCMDID].Text.Trim() == 
											Convert.ToString(ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Item(y).FVSCommandId).Trim() &&
											this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_CYCLE].Text.Trim() == 
											ReferenceFormRxPackageItem.ReferenceRxPackageItem.SimulationYear2Fvs.Trim())
											break;




									}
								}
								if (z>this.lvRxPackageFVSCmd.Items.Count-1)
								{
									this.AddItemToList(
										ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Item(y),
										ReferenceFormRxPackageItem.ReferenceRxPackageItem.SimulationYear2Fvs);
								}
								
							}
						}
						break;
					}
				}
				//
				//3rd simulation
				//
				for (x=0;x<=this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Count - 1;x++)
				{
					ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).FvsCycleList="";

					if (ReferenceFormRxPackageItem.ReferenceRxPackageItem.SimulationYear3Rx.Trim()==
						ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim())
					{
						if (ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection != null &&
							ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Count > 0)
						{
							for (y=0;y<=ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Count-1;y++)
							{
								//
								//make sure the rx,fvscmd and fvscmdid are not added already
								//
								for (z=0;z<=this.lvRxPackageFVSCmd.Items.Count-1;z++)
								{
									if (this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_RX].Text.Trim().Length > 0)
									{
										if (this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_RX].Text.Trim() ==
											ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim() &&
											this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_FVSCMD].Text.Trim().ToUpper() ==
											ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Item(y).FVSCommand.Trim().ToUpper() &&
											this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_FVSCMDID].Text.Trim() == 
											Convert.ToString(ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Item(y).FVSCommandId).Trim() &&
											this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_CYCLE].Text.Trim() == 
											ReferenceFormRxPackageItem.ReferenceRxPackageItem.SimulationYear3Fvs.Trim())
											break;




									}
								}
								if (z>this.lvRxPackageFVSCmd.Items.Count-1)
								{
									this.AddItemToList(
										ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Item(y),
										ReferenceFormRxPackageItem.ReferenceRxPackageItem.SimulationYear3Fvs);
								}
							}
						}
						break;
					}
				}
				//
				//4th simulation
				//
				for (x=0;x<=this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Count - 1;x++)
				{
					ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).FvsCycleList="";

					if (ReferenceFormRxPackageItem.ReferenceRxPackageItem.SimulationYear4Rx.Trim()==
						ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim())
					{
						if (ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection != null &&
							ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Count > 0)
						{
							for (y=0;y<=ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Count-1;y++)
							{
								//
								//make sure the rx,fvscmd and fvscmdid are not added already
								//
								for (z=0;z<=this.lvRxPackageFVSCmd.Items.Count-1;z++)
								{
									if (this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_RX].Text.Trim().Length > 0)
									{
										if (this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_RX].Text.Trim() ==
											ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim() &&
											this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_FVSCMD].Text.Trim().ToUpper() ==
											ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Item(y).FVSCommand.Trim().ToUpper() &&
											this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_FVSCMDID].Text.Trim() == 
											Convert.ToString(ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Item(y).FVSCommandId).Trim() &&
											this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_CYCLE].Text.Trim() == 
											ReferenceFormRxPackageItem.ReferenceRxPackageItem.SimulationYear4Fvs.Trim())
											break;




									}
								}
								if (z>this.lvRxPackageFVSCmd.Items.Count-1)
								{
									this.AddItemToList(
										ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceFvsCommandsCollection.Item(y),
										ReferenceFormRxPackageItem.ReferenceRxPackageItem.SimulationYear4Fvs);
								}
							}
						}
						break;
					}
				}
			
			


		}
		
		
		public void savevalues()
		{	
			 int x;

			
				if (ReferenceFormRxPackageItem.m_oRxPackageItem.m_oFvsCommandItem_Collection1 != null)
				{
			
					//make sure each command is assigned the rxid
					for (x=0;x<=ReferenceFormRxPackageItem.m_oRxPackageItem.m_oFvsCommandItem_Collection1.Count-1;x++)
					{
						if (ReferenceFormRxPackageItem.m_oRxPackageItem.m_oFvsCommandItem_Collection1.Item(x).RxPackageId.Trim().Length ==0)
						{
						   ReferenceFormRxPackageItem.m_oRxPackageItem.m_oFvsCommandItem_Collection1.Item(x).RxPackageId=this.ReferenceFormRxPackageItem.m_oRxPackageItem.RxPackageId;
						
						}
					}
				}
				
			this.ListViewOrder(this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oFvsCommandItem_Collection1);

			if (this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1 == null)
			{
				this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1 = new RxPackageCombinedFVSCommandsItem_Collection();
			}

			for (x=this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1.Count-1;x>=0;x--)
				this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1.Remove(x);

			
			for (x=0;x<=this.lvRxPackageFVSCmd.Items.Count-1;x++)
			{
				FIA_Biosum_Manager.RxPackageCombinedFVSCommandsItem oItem = new RxPackageCombinedFVSCommandsItem();
				oItem.Index = x;
				oItem.RxPackageId = this.ReferenceFormRxPackageItem.m_oRxPackageItem.RxPackageId;
				oItem.RxId = this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_RX].Text;
				oItem.FVSCommand = this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_FVSCMD].Text;
				oItem.FVSCommandId = Convert.ToByte(this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_FVSCMDID].Text);
				oItem.FVSCycle = this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text.Trim();
				this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1.Add(oItem);
			}
			this.ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceRxPackageCombinedFVSCommandsItemCollection = 
				this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1;

			this.ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection = 
				this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oFvsCommandItem_Collection1;
			
			
		}

		private void lvRxPackageFVSCmd_SelectedIndexChanged(object sender, System.EventArgs e)
		{

			
			
			if (this.lvRxPackageFVSCmd.SelectedItems.Count > 0)
			{
				m_intCurrSelect = lvRxPackageFVSCmd.SelectedItems[0].Index;
				this.m_oLvAlternateColors.DelegateListViewItem(lvRxPackageFVSCmd.SelectedItems[0]);
				if (this.lvRxPackageFVSCmd.SelectedItems[0].SubItems[COLUMN_CYCLE].Text.Trim().Length > 0)
				{
					if (this.ReferenceFormRxPackageItem.TabPageHasFocus(2)==true)
					{
                        ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_NEW] = true;
                        ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_DELETE] = false;
                        ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_CLEARALL] = true;
                        ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_EDIT] = false;
                        ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_OPEN] = true;
                        ReferenceFormRxPackageItem.SetToolBarButtonsEnabled(frmRxPackageItem.UC_FVSCMD);

					}
					this.btnDown.Enabled=false;
					this.btnUp.Enabled=false;
				}
				else
				{
					if (this.ReferenceFormRxPackageItem.TabPageHasFocus(2)==true)
					{
                        ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_NEW] = true;
                        ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_DELETE] = true;
                        ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_CLEARALL] = true;
                        ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_EDIT] = true;
                        ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_OPEN] = true;
                        ReferenceFormRxPackageItem.SetToolBarButtonsEnabled(frmRxPackageItem.UC_FVSCMD);

					}
					this.btnDown.Enabled=true;
					this.btnUp.Enabled=true;
				}
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
			string strFvsCmd="";
			FIA_Biosum_Manager.frmRxItemFvsCmdItem frmFvsCmdItem1 = new frmRxItemFvsCmdItem();
			FIA_Biosum_Manager.RxItemFvsCommandItem oRxItemFvsCmdItem=null;
			FIA_Biosum_Manager.RxPackageItemFvsCommandItem oRxPackageItemFvsCmdItem=null;
			frmFvsCmdItem1.MaximizeBox = true;
			frmFvsCmdItem1.BackColor = System.Drawing.SystemColors.Control;
			frmFvsCmdItem1.Text = "FVS: FVS Command Item (Edit)";
			

			//frmFvsCmdItem1.Initialize_Rx_User_Control();

			

			
			
			//frmFvsCmdItem1.uc_rx_edit1.m_oResizeForm.ScrollBarParentControl=frmFvsCmdItem1.uc_rx_edit1.ParentForm;
			


			
			//frmFvsCmdItem1.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmFvsCmdItem1.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			//find the current rxid
			//for (x=0;x<=this.m_oRxItem_Collection.Count-1;x++)
			//{
			//	if (this.m_oRxItem_Collection.Item(x).RxId.Trim() == 
			//		this.lstRx.SelectedItems[0].SubItems[COLUMN_RX].Text.Trim())
			//	{
			//		frmFvsCmdItem1.ReferenceRxItem = this.m_oRxItem_Collection.Item(x);
			//		break;
			//	}
			//}
			frmFvsCmdItem1.ReferenceUserControlPackageFvsCmdList=this;
			frmFvsCmdItem1.uc_rx_fvscmd_edit1.RxPackageEdit=true;
			
			frmFvsCmdItem1.m_strAction="edit";
			oRxPackageItemFvsCmdItem = this.ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(this.lvRxPackageFVSCmd.SelectedItems[0].Index);
			strFvsCmd = oRxPackageItemFvsCmdItem.FVSCommand;
			frmFvsCmdItem1.loadvalues(oRxPackageItemFvsCmdItem);
			
			System.Windows.Forms.DialogResult result = frmFvsCmdItem1.ShowDialog();
			if (result==System.Windows.Forms.DialogResult.OK)
			{
				if (oRxPackageItemFvsCmdItem.FVSCommand.Trim().ToUpper() != strFvsCmd.Trim().ToUpper())
				{
					RxTools oRxTools = new RxTools();
				    oRxPackageItemFvsCmdItem.FVSCommandId=oRxTools.AssignFvsCommandId(this.ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection,oRxPackageItemFvsCmdItem.FVSCommand);
					oRxTools=null;

				}
				
				oRxPackageItemFvsCmdItem.CopyProperties(frmFvsCmdItem1.uc_rx_fvscmd_edit1.m_oRxPackageItemFvsCommandItem,oRxPackageItemFvsCmdItem);
				UpdateListViewRxPackageItem(oRxPackageItemFvsCmdItem);
				
			
			}
		}
		private void UpdateListViewRxPackageItem(FIA_Biosum_Manager.RxPackageItemFvsCommandItem p_oRxPackageItemFvsCommandItem)
		{
			if (lvRxPackageFVSCmd.SelectedItems.Count==0) return;
			int x=lvRxPackageFVSCmd.SelectedItems[0].Index;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_PACKAGE].Text=p_oRxPackageItemFvsCommandItem.RxPackageId;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_FVSCMD].Text=p_oRxPackageItemFvsCommandItem.FVSCommand;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_P1].Text=p_oRxPackageItemFvsCommandItem.Parameter1;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_P2].Text=p_oRxPackageItemFvsCommandItem.Parameter2;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_P3].Text=p_oRxPackageItemFvsCommandItem.Parameter3;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_P4].Text=p_oRxPackageItemFvsCommandItem.Parameter4;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_P5].Text=p_oRxPackageItemFvsCommandItem.Parameter5;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_P6].Text=p_oRxPackageItemFvsCommandItem.Parameter6;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_P7].Text=p_oRxPackageItemFvsCommandItem.Parameter7;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_OTHER].Text=p_oRxPackageItemFvsCommandItem.Other;
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_FVSCMDID].Text = Convert.ToString(p_oRxPackageItemFvsCommandItem.FVSCommandId);
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
			this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_FVSCMDID].Text = Convert.ToString(p_oRxItemFvsCommandItem.FVSCommandId);
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
			    this.lvRxPackageFVSCmd.Items[p_intRow].SubItems[COLUMN_FVSCMDID].Text = Convert.ToString(p_oRxItemFvsCommandItem.FVSCommandId);
			
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

			for (y=0;y<=ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Count-1;y++)
			{
				if (ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(y).Delete==false)
				{
					
						if (this.lvRxPackageFVSCmd.Items[index].SubItems[COLUMN_PACKAGE].Text.Trim()==
							ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(y).RxPackageId.Trim() && 
							this.lvRxPackageFVSCmd.Items[index].SubItems[COLUMN_FVSCMD].Text.ToUpper().Trim() == 
							ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(y).FVSCommand.Trim().ToUpper() &&
							this.lvRxPackageFVSCmd.Items[index].SubItems[COLUMN_FVSCMDID].Text.Trim()==
							Convert.ToString(ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(y).FVSCommandId).Trim())
						{
							ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(y).Delete=true;
							ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(y).Index=-1;
							break;
						}
					
				}
			}


			
			/*
			
			//locate the current property associated with the listview
			for (x=0;x<=ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Count-1;x++)
			{
				if (ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(x).Index==index)
				{
					//ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Remove(x);
					ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(x).Delete=true;
					ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(x).Index=-1;
					//subtract 1 from the index of each item below the one we just removed
					for (y=0;y<=ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Count-1;y++)
					{
						if (ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(y).Delete==false)
						{
							if (ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(y).Index > index)
							{
								ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(y).Index = ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Item(y).Index - 1;
							}
						}
					}
					break;
				}
			}
			*/
			/**********************************************
			 **remove the ONE that is selected
			 **********************************************/
			if (index == 0 && lvRxPackageFVSCmd.Items.Count==1) 
			{
				lvRxPackageFVSCmd.Items.Clear();
			}
			else 
			{
					
				//*see if were at the top of the list
				if (index == 0 && lvRxPackageFVSCmd.Items.Count > 2) 
				{
					intSelect=0;
				}
				else 
				{
					//*see if were at the bottom
					if (index+1==lvRxPackageFVSCmd.Items.Count) 
					{
						this.m_intCurrSelect=index-1;
						intSelect=index-1;
					}
					else
					{
						intSelect=index;
					}
				}
				int intSelected = this.lvRxPackageFVSCmd.SelectedItems[0].Index;
				this.m_oLvAlternateColors.m_oRowCollection.Remove(intSelected);
				this.m_oLvAlternateColors.m_intSelectedRow=-1;

				lvRxPackageFVSCmd.Items.Remove(lvRxPackageFVSCmd.Items[index]);
			}

			if (lvRxPackageFVSCmd.Items.Count==0)
			{
                ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_NEW] = true;
                ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_DELETE] = false;
                ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_CLEARALL] = false;
                ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_EDIT] = false;
                ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_OPEN] = true;
                ReferenceFormRxPackageItem.SetToolBarButtonsEnabled(frmRxPackageItem.UC_FVSCMD);

			}
			else
			{
				this.lvRxPackageFVSCmd.Items[this.m_intCurrSelect].Selected =true;

			}
								
		}
		
		public void RemoveRxItemsFromList(string p_strRx, string p_strFvsCycle)
		{
			int x;
			int y;
            
			for (x=this.lvRxPackageFVSCmd.Items.Count-1;x>=0;x--)
			{

				if (this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_RX].Text.Trim()==p_strRx.Trim() &&
					this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text.Trim()== p_strFvsCycle)
					this.lvRxPackageFVSCmd.Items.Remove(lvRxPackageFVSCmd.Items[x]);
                    

			}
           

		}
		public void AddItem()
		{
			
			int x;
			
			FIA_Biosum_Manager.frmRxItemFvsCmdItem frmFvsCmdItem1 = new frmRxItemFvsCmdItem();
			//FIA_Biosum_Manager.RxItemFvsCommandItem oFvsCmdItem=null;
			FIA_Biosum_Manager.RxPackageItemFvsCommandItem oFvsCmdItem = null;
			frmFvsCmdItem1.MaximizeBox = true;
			frmFvsCmdItem1.BackColor = System.Drawing.SystemColors.Control;
			frmFvsCmdItem1.Text = "FVS: FVS Command Item (New)";
			

			frmFvsCmdItem1.uc_rx_fvscmd_edit1.RxPackageEdit=true;

			frmFvsCmdItem1.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			
			frmFvsCmdItem1.ReferenceUserControlPackageFvsCmdList=this;
			
			
			frmFvsCmdItem1.m_strAction="new";

			
				
					
				
			
			
			frmFvsCmdItem1.loadvalues(oFvsCmdItem);
			System.Windows.Forms.DialogResult result = frmFvsCmdItem1.ShowDialog();
			if (result==System.Windows.Forms.DialogResult.OK)
			{
				frmFvsCmdItem1.uc_rx_fvscmd_edit1.savevalues();
				//frmFvsCmdItem1.uc_rx_fvscmd_edit1.m_oFvsCmdItem.this.ReferenceFormRxPackageItem.m_oRxPackageItem.RxPackageId;
				frmFvsCmdItem1.uc_rx_fvscmd_edit1.m_oRxPackageItemFvsCommandItem.RxPackageId = this.ReferenceFormRxPackageItem.m_oRxPackageItem.RxPackageId;
				frmFvsCmdItem1.uc_rx_fvscmd_edit1.m_oRxPackageItemFvsCommandItem.Index = this.lvRxPackageFVSCmd.Items.Count;
				if (this.ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection==null)
				{
					this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oFvsCommandItem_Collection1= new RxPackageItemFvsCommandItem_Collection();
					this.ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection=this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oFvsCommandItem_Collection1;
				}
				RxTools oRxTools = new RxTools();
				frmFvsCmdItem1.uc_rx_fvscmd_edit1.m_oRxPackageItemFvsCommandItem.FVSCommandId=
					oRxTools.AssignFvsCommandId(
					   ReferenceFormRxPackageItem.m_oRxPackageItem.m_oFvsCommandItem_Collection1, 
					   frmFvsCmdItem1.uc_rx_fvscmd_edit1.m_oRxPackageItemFvsCommandItem.FVSCommand);

				oRxTools=null;
				this.ReferenceFormRxPackageItem.m_oRxPackageItem.ReferenceFvsCommandsCollection.Add(frmFvsCmdItem1.uc_rx_fvscmd_edit1.m_oRxPackageItemFvsCommandItem);
			    //ReferenceFormRxItem.m_oRxItem.ReferenceFvsCommandsCollection.Add(frmFvsCmdItem1.uc_rx_fvscmd_edit1.m_oRxItemFvsCmdItem);

				AddItemToList(frmFvsCmdItem1.uc_rx_fvscmd_edit1.m_oRxPackageItemFvsCommandItem);
				
			}
           
			
		}
		private void btnEdit_Click(object sender, System.EventArgs e)
		{
			EditItem();
		}
		private void AddItemToList(FIA_Biosum_Manager.RxItemFvsCommandItem oItem,string p_strCycle)
		{
			lvRxPackageFVSCmd.ListViewItemSorter = null;
			this.lvRxPackageFVSCmd.Items.Add("");
			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].UseItemStyleForSubItems=false;
			for (int z=1;z<=this.lvRxPackageFVSCmd.Columns.Count-1;z++)
			{
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems.Add(" ");
			}
			
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_RX].Text=oItem.RxId;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_CYCLE].Text=p_strCycle;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_FVSCMD].Text=oItem.FVSCommand;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P1].Text=oItem.Parameter1;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P2].Text=oItem.Parameter2;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P3].Text=oItem.Parameter3;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P4].Text=oItem.Parameter4;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P5].Text=oItem.Parameter5;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P6].Text=oItem.Parameter6;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P7].Text=oItem.Parameter7;
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_OTHER].Text=oItem.Other;
			    this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_FVSCMDID].Text = Convert.ToString(oItem.FVSCommandId);
		

									
			this.m_oLvAlternateColors.AddRow();
			this.m_oLvAlternateColors.AddColumns(lvRxPackageFVSCmd.Items.Count-1,this.lvRxPackageFVSCmd.Columns.Count);

			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].Selected=true;

            lvwColumnSorter = new ListViewColumnSorter();
            this.lvRxPackageFVSCmd.ListViewItemSorter = lvwColumnSorter;

			//move it in front of the other later cycle's
			//int intMoveUpCount=0;
			//for (x=this.lvRxPackageFVSCmd.Items.Count-1;x>=0;x--)
			//{
			//	if (this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text.Trim().Length > 0)
			//	{
			//		if (Convert.ToInt32(this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text) > 
			//			Convert.ToInt32(p_strCycle))
			//		{
			//			intMoveUpCount = lvRxPackageFVSCmd.Items.Count - x - 1;
			//		}
			//	}
			//}
			//if (intMoveUpCount > 0)
			//{
			//	for (x=1; x<= intMoveUpCount;x++)
			//	{
			//		MoveUp();
			//	}
			//}

            if (this.ReferenceFormRxPackageItem.TabPageHasFocus(2) == true)
            {
                ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_NEW] = true;
                ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_DELETE] = true;
                ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_CLEARALL] = true;
                ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_EDIT] = true;
                ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_OPEN] = true;
                ReferenceFormRxPackageItem.SetToolBarButtonsEnabled(frmRxPackageItem.UC_FVSCMD);

            }

           

			
		}
		public void UpdateRxFvsCycleItem(string p_strRx,string p_strCurFvsCycle,string p_strNewFvsCycle)
		{
			int x,y;
			for (x=this.lvRxPackageFVSCmd.Items.Count-1;x>=0;x--)
			{
				if (this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_RX].Text.Trim()==p_strRx.Trim() &&
					this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text.Trim()== p_strCurFvsCycle)
				{
					this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text = p_strNewFvsCycle;
				}
			}
		}
		public void AddRxItemsToList(string p_strRx,string p_strFvsCycle)
		{
			bool bFound=false;
			int x,y;
			for (x=this.lvRxPackageFVSCmd.Items.Count-1;x>=0;x--)
			{

				if (this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_RX].Text.Trim()==p_strRx.Trim() &&
					this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text.Trim()== p_strFvsCycle)
				{
					bFound=true;
					break;
				}

			}
			if (bFound==false)
			{
				//remove any treatments that have the same cycle as our current treatment
				for (x=this.lvRxPackageFVSCmd.Items.Count-1;x>=0;x--)
				{
					if (this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_RX].Text.Trim()!=p_strRx.Trim() &&
						this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text.Trim()== p_strFvsCycle)
					{
						this.lvRxPackageFVSCmd.Items.Remove(lvRxPackageFVSCmd.Items[x]);
					}
				}
				for (x=0;x<=this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Count-1;x++)
				{
					if (this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim()==
						p_strRx.Trim())
					{
						for (y=0;y<=ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).m_oFvsCommandItem_Collection1.Count-1;y++)
						{
							this.AddItemToList(ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).m_oFvsCommandItem_Collection1.Item(y),p_strFvsCycle);
						}						
					
					}
				}
			}
			//check to see if they need to be reordered
			//just concerned with rx items where later fvs cycles are before earlier fvs cycles
			int intFvsCycle=-1;
			for (x=0;x<=this.lvRxPackageFVSCmd.Items.Count-1;x++)
			{
				if (intFvsCycle==-1 && this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text.Trim().Length > 0)
				{
					intFvsCycle=Convert.ToInt32(lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text);
				}
				if (lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text.Trim().Length > 0)
				{
					if (Convert.ToInt32(lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text) > intFvsCycle)
					{
						ReOrderList();
						break;
					}
				}
				
			}
		}
		private void ReOrderList()
		{
           
			System.Windows.Forms.ListView oLv;
			int x,y,z;
			string strCurFvsCycle="";
			int intSet=0;
			int [] intSetArray = new int[this.lvRxPackageFVSCmd.Items.Count];
			//put the current list into sets
			for (x=0;x<=this.lvRxPackageFVSCmd.Items.Count-1;x++)
			{
				if (x==0)
				{
					intSet=0;
					intSetArray[x]=0;
					strCurFvsCycle=lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text;
				}
				else if (strCurFvsCycle != lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text)
				{
					intSet++;
					intSetArray[x]=intSet;
					strCurFvsCycle=lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text;
				}
				else
				{
					intSetArray[x]=intSet;
				}
			}
			//get the min and max cycle year values
			int intMin=1000;
			int intMax=-1;
			for (x=0;x<=this.lvRxPackageFVSCmd.Items.Count-1;x++)
			{
				if (this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text.Trim().Length!=0)
				{
					if (Convert.ToInt32(lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text) <
						intMin)
					{
						intMin = Convert.ToInt32(lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text);
					}
					else if (Convert.ToInt32(lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text) >
						intMax)
					{
						intMax = Convert.ToInt32(lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_CYCLE].Text);
					}
				}
			}
			//reorder the set 
			int intOrder=-1;
			string strCurrentFvsCycle="";
			for (x=intMin;x<=intMax;x++)
			{
				for (y=0;y<=this.lvRxPackageFVSCmd.Items.Count-1;y++)
				{
					//bypass package command
					if (this.lvRxPackageFVSCmd.Items[y].SubItems[COLUMN_CYCLE].Text.Trim().Length!=0)
					{
						//check if the cycle exists
						if (Convert.ToInt32(lvRxPackageFVSCmd.Items[y].SubItems[COLUMN_CYCLE].Text) ==x)
						{
							    //increment the order index
								intOrder++;				
							    //check to see if package command already equals the order index
								for (z=0;z<=this.lvRxPackageFVSCmd.Items.Count-1;z++)
								{
									if (this.lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_CYCLE].Text.Trim().Length==0 && 
										intSetArray[z]==intOrder)
									{
										//found package command that already equals order index so increment order index by 1
										intOrder++;
										break;
									}
								}
							
							strCurrentFvsCycle=lvRxPackageFVSCmd.Items[y].SubItems[COLUMN_CYCLE].Text;
							for (z=y;z<=this.lvRxPackageFVSCmd.Items.Count-1;z++)
							{
								//everything with the current fvs cycle assign the new order index
								if (strCurrentFvsCycle==lvRxPackageFVSCmd.Items[z].SubItems[COLUMN_CYCLE].Text)
								{
									intSetArray[z]=intOrder;
								}
								else break;
							}
							break;
							
						}

					}
				}
			}

			//copy the new order to the work listview object
			oLv = new ListView();

			for (x=0;x<=intSet;x++)
			{
				for (y=0;y<=intSetArray.Length - 1;y++)
				{
					if (intSetArray[y] == x)
					{
						oLv.Items.Add("");
						for (z=1;z<=this.lvRxPackageFVSCmd.Columns.Count-1;z++)
						{
							oLv.Items[oLv.Items.Count-1].SubItems.Add(this.lvRxPackageFVSCmd.Items[y].SubItems[z].Text);
						}


					}
				}
			}
			//copy the work listview values to the destination listview
			for (x=0;x<=this.lvRxPackageFVSCmd.Items.Count-1;x++)
			{
				for (z=1;z<=this.lvRxPackageFVSCmd.Columns.Count-1;z++)
				{
					lvRxPackageFVSCmd.Items[x].SubItems[z].Text = oLv.Items[x].SubItems[z].Text;
				}
			}

			oLv.Dispose();
			
			

	

			
		
		}
		private void AddItemToList(FIA_Biosum_Manager.RxPackageItemFvsCommandItem oItem)
		{
            lvRxPackageFVSCmd.ListViewItemSorter = null;
			this.lvRxPackageFVSCmd.Items.Add("");
			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].UseItemStyleForSubItems=false;
			for (int z=1;z<=this.lvRxPackageFVSCmd.Columns.Count-1;z++)
			{
				this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems.Add(" ");
			}
			
			
			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_PACKAGE].Text=
				      ReferenceFormRxPackageItem.uc_rx_package_edit1.RxPackageId;
			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_FVSCMD].Text=oItem.FVSCommand;
			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P1].Text=oItem.Parameter1;
			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P2].Text=oItem.Parameter2;
			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P3].Text=oItem.Parameter3;
			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P4].Text=oItem.Parameter4;
			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P5].Text=oItem.Parameter5;
			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P6].Text=oItem.Parameter6;
			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_P7].Text=oItem.Parameter7;
			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_OTHER].Text=oItem.Other;
			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].SubItems[COLUMN_FVSCMDID].Text = Convert.ToString(oItem.FVSCommandId);
			
									
			this.m_oLvAlternateColors.AddRow();
			this.m_oLvAlternateColors.AddColumns(lvRxPackageFVSCmd.Items.Count-1,this.lvRxPackageFVSCmd.Columns.Count);

			

			this.lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items.Count-1].Selected=true;

            lvwColumnSorter = new ListViewColumnSorter();
            this.lvRxPackageFVSCmd.ListViewItemSorter = lvwColumnSorter;

            if (this.ReferenceFormRxPackageItem.TabPageHasFocus(2) == true)
            {
                ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_NEW] = true;
                ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_DELETE] = true;
                ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_CLEARALL] = true;
                ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_EDIT] = true;
                ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_OPEN] = true;
                ReferenceFormRxPackageItem.SetToolBarButtonsEnabled(frmRxPackageItem.UC_FVSCMD);

            }
		}
		/// <summary>
		/// Assign an index value equal to the fvs commands location in the listview
		/// </summary>
		/// <param name="p_oFvsCollection"></param>
		public void ListViewOrder(FIA_Biosum_Manager.RxPackageItemFvsCommandItem_Collection p_oFvsCollection)
		{
			if (p_oFvsCollection == null || p_oFvsCollection.Count==0) return;
			FIA_Biosum_Manager.RxPackageItemFvsCommandItem oItem;
			int intMin=10,intMax=0;
			int x,y;
			for (y=0;y<=p_oFvsCollection.Count-1;y++)
			{
				if (p_oFvsCollection.Item(y).Delete==false)
				{
					for (x=0;x<=this.lvRxPackageFVSCmd.Items.Count-1;x++)
					{
						if (this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_PACKAGE].Text.Trim()==
							p_oFvsCollection.Item(y).RxPackageId.Trim() && 
							this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_FVSCMD].Text.ToUpper().Trim() == 
							p_oFvsCollection.Item(y).FVSCommand.Trim().ToUpper() &&
							this.lvRxPackageFVSCmd.Items[x].SubItems[COLUMN_FVSCMDID].Text.Trim()==
							Convert.ToString(p_oFvsCollection.Item(y).FVSCommandId).Trim())
						{
							if (x > intMax) intMax=x;
							if (x < intMin) intMin=x; 
							p_oFvsCollection.Item(y).ListViewIndex=x;
							
						}
					}
				}
			}
			FIA_Biosum_Manager.RxPackageItemFvsCommandItem_Collection oFvsCollection = new RxPackageItemFvsCommandItem_Collection();
			for (x=intMin;x<=intMax;x++)
			{
				for (y=0;y<=p_oFvsCollection.Count-1;y++)
				{
					if (x==p_oFvsCollection.Item(y).ListViewIndex)
					{
						oItem = new RxPackageItemFvsCommandItem();
						
						oItem.CopyProperties(p_oFvsCollection.Item(y),oItem);
						oFvsCollection.Add(oItem);
					}

				}
			}

			for (x=p_oFvsCollection.Count-1;x>=0;x--)
			{
				p_oFvsCollection.Remove(x);
			}

			for (x=0;x<=oFvsCollection.Count-1;x++)
			{
				oItem = new RxPackageItemFvsCommandItem();
				oItem.CopyProperties(oFvsCollection.Item(x),oItem);
				p_oFvsCollection.Add(oItem);
			}

			
		}
		
		public void RemoveAllItems()
		{
			if (lvRxPackageFVSCmd.SelectedItems.Count==0) return;
			if (this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oFvsCommandItem_Collection1==null) return;

			DialogResult result = MessageBox.Show("This will clear only Package commands.  Rx associated commands will not be removed. Do you still wish to remove the Package commands?(Y/N)","FIA Biosum",System.Windows.Forms.MessageBoxButtons.YesNo,System.Windows.Forms.MessageBoxIcon.Question);
			if (result==System.Windows.Forms.DialogResult.Yes)
			{
				
				for (int x=0;x<=this.ReferenceFormRxPackageItem.m_oRxPackageItem.m_oFvsCommandItem_Collection1.Count-1;x++)
				{
					ReferenceFormRxPackageItem.m_oRxPackageItem.m_oFvsCommandItem_Collection1.Item(x).Delete=true;
					ReferenceFormRxPackageItem.m_oRxPackageItem.m_oFvsCommandItem_Collection1.Item(x).Index=-1;
					for (int y=0;y<=this.lvRxPackageFVSCmd.Items.Count-1;y++)
					{
						if (lvRxPackageFVSCmd.Items[y].SubItems[COLUMN_PACKAGE].Text.Trim().Length > 0 &&
							lvRxPackageFVSCmd.Items[y].SubItems[COLUMN_RX].Text.Trim().Length == 0)
						{
							this.lvRxPackageFVSCmd.Items.Remove(lvRxPackageFVSCmd.Items[y]);
							
							
						}
					}
					

				}
				


			}

		}
			
		public FIA_Biosum_Manager.frmRxPackageItem ReferenceFormRxPackageItem
		{
			get {return this._frmRxPackageItem;}
			set {this._frmRxPackageItem=value;}

		}

        private void lvRxPackageFVSCmd_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            int x, y;

            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            lvRxPackageFVSCmd.Sort();
            //reinitialize the alternate row colors
            for (x = 0; x <= lvRxPackageFVSCmd.Items.Count - 1; x++)
            {
                this.lvRxPackageFVSCmd.Items[x].Selected = false;
                for (y = 0; y <= lvRxPackageFVSCmd.Columns.Count - 1; y++)
                {
                    this.m_oLvAlternateColors.ListViewSubItem(lvRxPackageFVSCmd.Items[x].Index, y, lvRxPackageFVSCmd.Items[lvRxPackageFVSCmd.Items[x].Index].SubItems[y], false);
                }
            }
            ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_NEW] = true;
            ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_DELETE] = false;
            ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_CLEARALL] = true;
            ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_EDIT] = false;
            ReferenceFormRxPackageItem.m_bToolBarButtonEnabled[frmRxPackageItem.UC_FVSCMD, frmRxPackageItem.BUTTON_OPEN] = true;
            ReferenceFormRxPackageItem.SetToolBarButtonsEnabled(frmRxPackageItem.UC_FVSCMD);
            
			
        }

	}

}
