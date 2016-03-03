using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_tree_species.
	/// </summary>
	public class uc_tree_species : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.GroupBox grpboxAudit;
		public System.Windows.Forms.ListView lstAudit;
		private System.Windows.Forms.Button btnHelp;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnCancel;
		public System.Windows.Forms.Button btnSave;
		private System.Windows.Forms.Button btnEdit;
		private System.Windows.Forms.Button btnNew;
		public System.Windows.Forms.Button btnDelete;
		private string m_strProjDir;
		private System.Windows.Forms.GroupBox grpBoxTreeSpc;
		public System.Windows.Forms.ListView lstTreeSpc;
		private FIA_Biosum_Manager.Datasource m_DataSource;
		private string m_strPlotTable;
		private string m_strTreeTable;
		private string m_strTreeSpcTable;
		private string m_strCondTable;
		private string m_strTempMDBFile;
		private FIA_Biosum_Manager.ado_data_access m_ado;
		private string m_strConn;
		private System.Windows.Forms.Button btnAuditFVSOutput;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_tree_species(string p_strProjDir)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.m_strProjDir = p_strProjDir;
			this.m_DataSource = new Datasource(p_strProjDir);
			this.m_strPlotTable = this.m_DataSource.getValidDataSourceTableName("PLOT");
			this.m_strTreeTable = this.m_DataSource.getValidDataSourceTableName("TREE");
			this.m_strTreeSpcTable = this.m_DataSource.getValidDataSourceTableName("TREE SPECIES");
			this.m_strCondTable = this.m_DataSource.getValidDataSourceTableName("CONDITION");
			this.m_strTempMDBFile = this.m_DataSource.CreateMDBAndTableDataSourceLinks();
			this.m_ado = new ado_data_access();
			this.m_strConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"","");

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
			this.btnHelp = new System.Windows.Forms.Button();
			this.btnClose = new System.Windows.Forms.Button();
			this.grpBoxTreeSpc = new System.Windows.Forms.GroupBox();
			this.btnDelete = new System.Windows.Forms.Button();
			this.lstTreeSpc = new System.Windows.Forms.ListView();
			this.btnNew = new System.Windows.Forms.Button();
			this.btnEdit = new System.Windows.Forms.Button();
			this.btnSave = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.grpboxAudit = new System.Windows.Forms.GroupBox();
			this.btnAuditFVSOutput = new System.Windows.Forms.Button();
			this.lstAudit = new System.Windows.Forms.ListView();
			this.lblTitle = new System.Windows.Forms.Label();
			this.groupBox1.SuspendLayout();
			this.grpBoxTreeSpc.SuspendLayout();
			this.grpboxAudit.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btnHelp);
			this.groupBox1.Controls.Add(this.btnClose);
			this.groupBox1.Controls.Add(this.grpBoxTreeSpc);
			this.groupBox1.Controls.Add(this.grpboxAudit);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(744, 600);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// btnHelp
			// 
			this.btnHelp.Location = new System.Drawing.Point(8, 560);
			this.btnHelp.Name = "btnHelp";
			this.btnHelp.Size = new System.Drawing.Size(96, 32);
			this.btnHelp.TabIndex = 45;
			this.btnHelp.Text = "Help";
			// 
			// btnClose
			// 
			this.btnClose.Location = new System.Drawing.Point(640, 560);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(96, 32);
			this.btnClose.TabIndex = 44;
			this.btnClose.Text = "Close";
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// grpBoxTreeSpc
			// 
			this.grpBoxTreeSpc.Controls.Add(this.btnDelete);
			this.grpBoxTreeSpc.Controls.Add(this.lstTreeSpc);
			this.grpBoxTreeSpc.Controls.Add(this.btnNew);
			this.grpBoxTreeSpc.Controls.Add(this.btnEdit);
			this.grpBoxTreeSpc.Controls.Add(this.btnSave);
			this.grpBoxTreeSpc.Controls.Add(this.btnCancel);
			this.grpBoxTreeSpc.Location = new System.Drawing.Point(16, 232);
			this.grpBoxTreeSpc.Name = "grpBoxTreeSpc";
			this.grpBoxTreeSpc.Size = new System.Drawing.Size(712, 320);
			this.grpBoxTreeSpc.TabIndex = 28;
			this.grpBoxTreeSpc.TabStop = false;
			this.grpBoxTreeSpc.Text = "Tree Species Table";
			// 
			// btnDelete
			// 
			this.btnDelete.Enabled = false;
			this.btnDelete.Location = new System.Drawing.Point(359, 243);
			this.btnDelete.Name = "btnDelete";
			this.btnDelete.Size = new System.Drawing.Size(64, 32);
			this.btnDelete.TabIndex = 51;
			this.btnDelete.Text = "Delete";
			// 
			// lstTreeSpc
			// 
			this.lstTreeSpc.FullRowSelect = true;
			this.lstTreeSpc.GridLines = true;
			this.lstTreeSpc.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
			this.lstTreeSpc.HideSelection = false;
			this.lstTreeSpc.Location = new System.Drawing.Point(8, 16);
			this.lstTreeSpc.MultiSelect = false;
			this.lstTreeSpc.Name = "lstTreeSpc";
			this.lstTreeSpc.Size = new System.Drawing.Size(696, 224);
			this.lstTreeSpc.TabIndex = 28;
			this.lstTreeSpc.View = System.Windows.Forms.View.Details;
			// 
			// btnNew
			// 
			this.btnNew.Location = new System.Drawing.Point(231, 243);
			this.btnNew.Name = "btnNew";
			this.btnNew.Size = new System.Drawing.Size(64, 32);
			this.btnNew.TabIndex = 46;
			this.btnNew.Text = "New";
			// 
			// btnEdit
			// 
			this.btnEdit.Location = new System.Drawing.Point(295, 243);
			this.btnEdit.Name = "btnEdit";
			this.btnEdit.Size = new System.Drawing.Size(64, 32);
			this.btnEdit.TabIndex = 47;
			this.btnEdit.Text = "Edit";
			// 
			// btnSave
			// 
			this.btnSave.Enabled = false;
			this.btnSave.Location = new System.Drawing.Point(423, 243);
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(64, 32);
			this.btnSave.TabIndex = 48;
			this.btnSave.Text = "Save";
			// 
			// btnCancel
			// 
			this.btnCancel.Location = new System.Drawing.Point(487, 243);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(64, 32);
			this.btnCancel.TabIndex = 49;
			this.btnCancel.Text = "Cancel";
			// 
			// grpboxAudit
			// 
			this.grpboxAudit.Controls.Add(this.btnAuditFVSOutput);
			this.grpboxAudit.Controls.Add(this.lstAudit);
			this.grpboxAudit.Location = new System.Drawing.Point(16, 48);
			this.grpboxAudit.Name = "grpboxAudit";
			this.grpboxAudit.Size = new System.Drawing.Size(712, 176);
			this.grpboxAudit.TabIndex = 27;
			this.grpboxAudit.TabStop = false;
			this.grpboxAudit.Text = "Audit Results";
			// 
			// btnAuditFVSOutput
			// 
			this.btnAuditFVSOutput.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnAuditFVSOutput.Location = new System.Drawing.Point(8, 138);
			this.btnAuditFVSOutput.Name = "btnAuditFVSOutput";
			this.btnAuditFVSOutput.Size = new System.Drawing.Size(664, 32);
			this.btnAuditFVSOutput.TabIndex = 30;
			this.btnAuditFVSOutput.Text = "FVS Output Tables: FVS Tree Species And Variant Combinations Not Found  In Tree S" +
				"pecies Table";
			// 
			// lstAudit
			// 
			this.lstAudit.FullRowSelect = true;
			this.lstAudit.GridLines = true;
			this.lstAudit.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
			this.lstAudit.HideSelection = false;
			this.lstAudit.Location = new System.Drawing.Point(6, 20);
			this.lstAudit.MultiSelect = false;
			this.lstAudit.Name = "lstAudit";
			this.lstAudit.Size = new System.Drawing.Size(696, 112);
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
			this.lblTitle.Size = new System.Drawing.Size(738, 32);
			this.lblTitle.TabIndex = 26;
			this.lblTitle.Text = "Tree Species";
			// 
			// uc_tree_species
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_tree_species";
			this.Size = new System.Drawing.Size(744, 600);
			this.Resize += new System.EventHandler(this.uc_tree_species_Resize);
			this.groupBox1.ResumeLayout(false);
			this.grpBoxTreeSpc.ResumeLayout(false);
			this.grpboxAudit.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void uc_tree_species_Resize(object sender, System.EventArgs e)
		{
			this.Resize_Tree_Species();
		}
		private void Resize_Tree_Species()
		{
			try
			{
				this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
				this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
				this.btnHelp.Top = this.btnClose.Top;
				this.grpboxAudit.Width = this.groupBox1.Width  - (this.grpboxAudit.Left * 2);
				this.grpBoxTreeSpc.Width = this.grpboxAudit.Width ;
				this.lstAudit.Width = this.grpboxAudit.Width - (this.lstAudit.Left * 2);
				this.lstTreeSpc.Width = this.grpBoxTreeSpc.Width - (this.lstTreeSpc.Left * 2);
				this.grpBoxTreeSpc.Height = this.btnClose.Top - this.grpBoxTreeSpc.Top - 5;
				this.btnNew.Top =this.grpBoxTreeSpc.Height - this.btnNew.Height - 2;
				this.btnEdit.Top = this.btnNew.Top;
				this.btnSave.Top = this.btnNew.Top;
				this.btnDelete.Top = this.btnNew.Top;
				this.btnCancel.Top = this.btnNew.Top;
				this.lstTreeSpc.Height = this.btnNew.Top - this.lstTreeSpc.Top - 2;
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

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.ParentForm.Close();
		}

		private void btnAuditPlotFVSVariants_Click(object sender, System.EventArgs e)
		{
		   this.MissingPlotVariants();
		}
		public void MissingPlotVariants()
		{
			int x;
			this.m_ado.m_intError=0;
			this.lstAudit.Clear();
			this.m_ado.m_strSQL = "select fvs_variant " + 
								  "from " + this.m_strPlotTable + " p " + 
								  "where not exists (select variant " + 
													"from " + this.m_strTreeSpcTable + " t " + 
													"where trim(p.fvs_variant)=trim(t.variant))";
            
			this.m_ado.SqlQueryReader(this.m_strConn,this.m_ado.m_strSQL);
			if (this.m_ado.m_intError==0)
			{
				if (this.m_ado.m_OleDbDataReader.HasRows)
				{
					for (x=0;x<=this.m_ado.m_OleDbDataReader.FieldCount-1;x++)
					{
						this.lstAudit.Columns.Add(this.m_ado.m_OleDbDataReader.GetName(x).ToString().Trim(), 80, HorizontalAlignment.Left);
					}

					while (this.m_ado.m_OleDbDataReader.Read())
					{
						if (this.m_ado.m_OleDbDataReader[0] != System.DBNull.Value)
						{
							this.lstAudit.Items.Add(this.m_ado.m_OleDbDataReader[0].ToString().Trim());
						}
						else
						{
							this.lstAudit.Items.Add(" ");
						}
						for (x=1;x<=this.m_ado.m_OleDbDataReader.FieldCount-1;x++)
						{
							this.lstAudit.Items[this.lstAudit.Items.Count-1].SubItems[x].Text = this.m_ado.m_OleDbDataReader[x].ToString().Trim();
						}

					}

					//for (x=0;x<=this.m_ado.m_OleDbDataReader.FieldCount-1;x++)
					//{
					//	this.lstAudit.Columns[x].Width = -1;
					//}
				}
				this.m_ado.m_OleDbDataReader.Close();
			}




		}
		public void loadvalues()
		{

			int x;
			this.m_ado.m_intError=0;
			this.lstTreeSpc.Clear();
			this.m_ado.m_strSQL = "select variant,spcd,fvs_species," + 
				                         "common_name,genus,species," + 
				                         "variety,subspecies,user_spc_group," + 
				                         "od_wgt,dry_to_green " + 
				                  "from " + this.m_strTreeSpcTable + " " + 
				                  "order by variant,spcd,fvs_species;";

			this.m_ado.SqlQueryReader(this.m_strConn,this.m_ado.m_strSQL);
			if (this.m_ado.m_intError==0)
			{
				if (this.m_ado.m_OleDbDataReader.HasRows)
				{
					for (x=0;x<=this.m_ado.m_OleDbDataReader.FieldCount-1;x++)
					{
						this.lstTreeSpc.Columns.Add(this.m_ado.m_OleDbDataReader.GetName(x).ToString().Trim(), 
							   100, HorizontalAlignment.Left);
					}

					while (this.m_ado.m_OleDbDataReader.Read())
					{
						if (this.m_ado.m_OleDbDataReader[0] != System.DBNull.Value)
						{
							this.lstTreeSpc.Items.Add(this.m_ado.m_OleDbDataReader[0].ToString().Trim());
						}
						else
						{
							this.lstTreeSpc.Items.Add(" ");
						}
						for (x=1;x<=this.m_ado.m_OleDbDataReader.FieldCount-1;x++)
						{
							this.lstTreeSpc.Items[this.lstTreeSpc.Items.Count-1].SubItems.Add(this.m_ado.m_OleDbDataReader[x].ToString().Trim());
						}

					}

					
				}
				this.m_ado.m_OleDbDataReader.Close();
			}

		}

		private void btnAuditPlotTreeVariantCombo_Click(object sender, System.EventArgs e)
		{
			this.MissingTreeSpcVariantCombo();
		
		}
		public void MissingTreeSpcVariantCombo()
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
				                  "where t.biosum_cond_id=c.biosum_cond_id " + 
				                  "order by p.fvs_variant,t.spcd;";


			this.m_ado.SqlQueryReader(this.m_strConn,this.m_ado.m_strSQL);
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
					this.m_ado.m_OleDbDataReader.Close();
					if (count > 0)
					{   
						int y=0;
						for (int x=0; x<=count-1;x++)
						{
							for (y=0; y<=this.lstTreeSpc.Items.Count-1;y++)
							{
								if (this.lstTreeSpc.Items[y].Text.Trim().ToUpper() == 
									strValues[x,1].ToString().Trim().ToUpper() &&
									this.lstTreeSpc.Items[y].SubItems[1].Text.Trim() == 
									strValues[x,0].ToString().Trim())
								{
									break;
								}
							}
							if (y > this.lstTreeSpc.Items.Count-1)
							{
								this.lstAudit.Items.Add(strValues[x,1].ToString().Trim());
								this.lstAudit.Items[this.lstAudit.Items.Count-1].SubItems.Add(strValues[x,0].ToString().Trim());
							}
						}
					}
				}
			}
		}
	}
}
