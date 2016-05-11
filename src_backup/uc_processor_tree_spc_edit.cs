using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_processor_tree_spc_edit.
	/// </summary>
	public class uc_processor_tree_spc_edit : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Label lblID;
		private System.Windows.Forms.ComboBox cmbVariant;
		private System.Windows.Forms.TextBox txtSpCd;
		private System.Windows.Forms.TextBox txtCommon;
		private int m_intError=0;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label lblUserTreeSpCd;
		private System.Windows.Forms.TextBox txtOvenDryWt;
		private System.Windows.Forms.TextBox txtGreenWt;
		private System.Windows.Forms.TextBox txtGenus;
		private System.Windows.Forms.TextBox txtSpecies;
		private System.Windows.Forms.TextBox txtVariety;
		private System.Windows.Forms.TextBox txtSubspecies;
		private FIA_Biosum_Manager.txtNumeric _txtOvenDryWt;
		private FIA_Biosum_Manager.txtNumeric _txtGreenWt;
		private System.Windows.Forms.TextBox txtConvertToSpCd;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.ComboBox cmbFvsSpCd;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label14;
		private string m_strFvsTreeSpcTable;
		private string m_strVariant;
		private ado_data_access m_ado;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.TextBox txtFvsCommonName;

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_processor_tree_spc_edit(ado_data_access p_ado,string p_strTreeSpcTable,string p_strFvsTreeSpcTable,string p_strVariant)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			_txtOvenDryWt = new txtNumeric(2,2);
			_txtGreenWt = new txtNumeric(2,2);
			this.groupBox1.Controls.Add(_txtOvenDryWt);
			this.groupBox1.Controls.Add(_txtGreenWt);

            _txtOvenDryWt.RightToLeft = System.Windows.Forms.RightToLeft.No;
			_txtGreenWt.RightToLeft = System.Windows.Forms.RightToLeft.No;
			_txtOvenDryWt.Size = this.txtOvenDryWt.Size;
			_txtOvenDryWt.Location = this.txtOvenDryWt.Location;
			this.txtOvenDryWt.Visible=false;
			_txtOvenDryWt.Visible=true;

			_txtGreenWt.Size = this.txtGreenWt.Size;
			_txtGreenWt.Location = this.txtGreenWt.Location;
			this.txtGreenWt.Visible=false;
			_txtGreenWt.Visible=true;
            m_ado = p_ado;
			if (p_strVariant.Trim().Length > 0)
			{
				p_ado.m_strSQL = "SELECT  spcd,fvs_species,common_name " + 
					"FROM " + p_strFvsTreeSpcTable + " " + 
					"WHERE LEN(TRIM(fvs_species)) > 0 AND " + 
					"LEN(TRIM(common_name)) > 0 AND " + 
					"TRIM(fvs_variant) = '" + p_strVariant.Trim() +  "' order by spcd;";
			
				p_ado.SqlQueryReader(p_ado.m_OleDbConnection,p_ado.m_OleDbTransaction,p_ado.m_strSQL);
				if (p_ado.m_OleDbDataReader.HasRows)
				{
					while (p_ado.m_OleDbDataReader.Read())
					{
						this.cmbFvsSpCd.Items.Add(Convert.ToString(p_ado.m_OleDbDataReader["fvs_species"]) + " - " + Convert.ToString(p_ado.m_OleDbDataReader["spcd"]) + " - " + Convert.ToString(p_ado.m_OleDbDataReader["common_name"]));

					}
				}
				p_ado.m_OleDbDataReader.Close();
			}
            this.m_strFvsTreeSpcTable = p_strFvsTreeSpcTable;
			this.m_strVariant = p_strVariant;

			


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
			this.txtFvsCommonName = new System.Windows.Forms.TextBox();
			this.label15 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.txtConvertToSpCd = new System.Windows.Forms.TextBox();
			this.label13 = new System.Windows.Forms.Label();
			this.cmbFvsSpCd = new System.Windows.Forms.ComboBox();
			this.label4 = new System.Windows.Forms.Label();
			this.txtSubspecies = new System.Windows.Forms.TextBox();
			this.txtVariety = new System.Windows.Forms.TextBox();
			this.txtSpecies = new System.Windows.Forms.TextBox();
			this.txtGenus = new System.Windows.Forms.TextBox();
			this.txtGreenWt = new System.Windows.Forms.TextBox();
			this.txtOvenDryWt = new System.Windows.Forms.TextBox();
			this.lblUserTreeSpCd = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.txtCommon = new System.Windows.Forms.TextBox();
			this.txtSpCd = new System.Windows.Forms.TextBox();
			this.cmbVariant = new System.Windows.Forms.ComboBox();
			this.lblID = new System.Windows.Forms.Label();
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnOK = new System.Windows.Forms.Button();
			this.label6 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.lblTitle = new System.Windows.Forms.Label();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.txtFvsCommonName);
			this.groupBox1.Controls.Add(this.label15);
			this.groupBox1.Controls.Add(this.label14);
			this.groupBox1.Controls.Add(this.txtConvertToSpCd);
			this.groupBox1.Controls.Add(this.label13);
			this.groupBox1.Controls.Add(this.cmbFvsSpCd);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Controls.Add(this.txtSubspecies);
			this.groupBox1.Controls.Add(this.txtVariety);
			this.groupBox1.Controls.Add(this.txtSpecies);
			this.groupBox1.Controls.Add(this.txtGenus);
			this.groupBox1.Controls.Add(this.txtGreenWt);
			this.groupBox1.Controls.Add(this.txtOvenDryWt);
			this.groupBox1.Controls.Add(this.lblUserTreeSpCd);
			this.groupBox1.Controls.Add(this.label12);
			this.groupBox1.Controls.Add(this.label11);
			this.groupBox1.Controls.Add(this.label10);
			this.groupBox1.Controls.Add(this.label9);
			this.groupBox1.Controls.Add(this.label8);
			this.groupBox1.Controls.Add(this.label7);
			this.groupBox1.Controls.Add(this.txtCommon);
			this.groupBox1.Controls.Add(this.txtSpCd);
			this.groupBox1.Controls.Add(this.cmbVariant);
			this.groupBox1.Controls.Add(this.lblID);
			this.groupBox1.Controls.Add(this.btnCancel);
			this.groupBox1.Controls.Add(this.btnOK);
			this.groupBox1.Controls.Add(this.label6);
			this.groupBox1.Controls.Add(this.label5);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.groupBox1.Size = new System.Drawing.Size(624, 608);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// txtFvsCommonName
			// 
			this.txtFvsCommonName.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtFvsCommonName.Location = new System.Drawing.Point(287, 520);
			this.txtFvsCommonName.MaxLength = 50;
			this.txtFvsCommonName.Name = "txtFvsCommonName";
			this.txtFvsCommonName.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtFvsCommonName.Size = new System.Drawing.Size(296, 23);
			this.txtFvsCommonName.TabIndex = 39;
			this.txtFvsCommonName.Text = "";
			// 
			// label15
			// 
			this.label15.Location = new System.Drawing.Point(32, 525);
			this.label15.Name = "label15";
			this.label15.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.label15.Size = new System.Drawing.Size(248, 16);
			this.label15.TabIndex = 38;
			this.label15.Text = "FVS Tree Species Common Name";
			// 
			// label14
			// 
			this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(24, 404);
			this.label14.Name = "label14";
			this.label14.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.label14.Size = new System.Drawing.Size(584, 37);
			this.label14.TabIndex = 36;
			this.label14.Text = "Optional: Enter an alternate FIA tree species code.  When the application creates" +
				" FVS input files, it will translate the tree species code to the tree species co" +
				"de defined below.";
			// 
			// txtConvertToSpCd
			// 
			this.txtConvertToSpCd.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtConvertToSpCd.Location = new System.Drawing.Point(288, 486);
			this.txtConvertToSpCd.MaxLength = 3;
			this.txtConvertToSpCd.Name = "txtConvertToSpCd";
			this.txtConvertToSpCd.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtConvertToSpCd.Size = new System.Drawing.Size(56, 23);
			this.txtConvertToSpCd.TabIndex = 34;
			this.txtConvertToSpCd.Text = "";
			// 
			// label13
			// 
			this.label13.Location = new System.Drawing.Point(32, 486);
			this.label13.Name = "label13";
			this.label13.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.label13.Size = new System.Drawing.Size(248, 16);
			this.label13.TabIndex = 35;
			this.label13.Text = "FVS Input FIA Tree Species Code";
			// 
			// cmbFvsSpCd
			// 
			this.cmbFvsSpCd.Location = new System.Drawing.Point(288, 448);
			this.cmbFvsSpCd.Name = "cmbFvsSpCd";
			this.cmbFvsSpCd.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.cmbFvsSpCd.Size = new System.Drawing.Size(296, 24);
			this.cmbFvsSpCd.TabIndex = 33;
			this.cmbFvsSpCd.SelectedIndexChanged += new System.EventHandler(this.cmbFvsSpCd_SelectedIndexChanged);
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(120, 448);
			this.label4.Name = "label4";
			this.label4.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.label4.Size = new System.Drawing.Size(160, 16);
			this.label4.TabIndex = 32;
			this.label4.Text = "FVS Tree Species Code";
			// 
			// txtSubspecies
			// 
			this.txtSubspecies.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSubspecies.Location = new System.Drawing.Point(280, 309);
			this.txtSubspecies.MaxLength = 50;
			this.txtSubspecies.Name = "txtSubspecies";
			this.txtSubspecies.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtSubspecies.Size = new System.Drawing.Size(296, 23);
			this.txtSubspecies.TabIndex = 28;
			this.txtSubspecies.Text = "";
			// 
			// txtVariety
			// 
			this.txtVariety.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVariety.Location = new System.Drawing.Point(280, 279);
			this.txtVariety.MaxLength = 50;
			this.txtVariety.Name = "txtVariety";
			this.txtVariety.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtVariety.Size = new System.Drawing.Size(296, 23);
			this.txtVariety.TabIndex = 27;
			this.txtVariety.Text = "";
			// 
			// txtSpecies
			// 
			this.txtSpecies.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSpecies.Location = new System.Drawing.Point(280, 248);
			this.txtSpecies.MaxLength = 50;
			this.txtSpecies.Name = "txtSpecies";
			this.txtSpecies.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtSpecies.Size = new System.Drawing.Size(296, 23);
			this.txtSpecies.TabIndex = 26;
			this.txtSpecies.Text = "";
			// 
			// txtGenus
			// 
			this.txtGenus.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtGenus.Location = new System.Drawing.Point(280, 184);
			this.txtGenus.MaxLength = 50;
			this.txtGenus.Name = "txtGenus";
			this.txtGenus.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtGenus.Size = new System.Drawing.Size(296, 23);
			this.txtGenus.TabIndex = 25;
			this.txtGenus.Text = "";
			// 
			// txtGreenWt
			// 
			this.txtGreenWt.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtGreenWt.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtGreenWt.Location = new System.Drawing.Point(280, 369);
			this.txtGreenWt.MaxLength = 2;
			this.txtGreenWt.Name = "txtGreenWt";
			this.txtGreenWt.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtGreenWt.Size = new System.Drawing.Size(112, 23);
			this.txtGreenWt.TabIndex = 24;
			this.txtGreenWt.Text = "";
			// 
			// txtOvenDryWt
			// 
			this.txtOvenDryWt.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtOvenDryWt.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtOvenDryWt.Location = new System.Drawing.Point(280, 339);
			this.txtOvenDryWt.MaxLength = 2;
			this.txtOvenDryWt.Name = "txtOvenDryWt";
			this.txtOvenDryWt.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtOvenDryWt.Size = new System.Drawing.Size(112, 23);
			this.txtOvenDryWt.TabIndex = 23;
			this.txtOvenDryWt.Text = "";
			// 
			// lblUserTreeSpCd
			// 
			this.lblUserTreeSpCd.BackColor = System.Drawing.Color.White;
			this.lblUserTreeSpCd.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUserTreeSpCd.Location = new System.Drawing.Point(280, 216);
			this.lblUserTreeSpCd.Name = "lblUserTreeSpCd";
			this.lblUserTreeSpCd.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.lblUserTreeSpCd.Size = new System.Drawing.Size(40, 24);
			this.lblUserTreeSpCd.TabIndex = 21;
			this.lblUserTreeSpCd.Text = "0";
			// 
			// label12
			// 
			this.label12.Location = new System.Drawing.Point(104, 309);
			this.label12.Name = "label12";
			this.label12.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.label12.Size = new System.Drawing.Size(160, 16);
			this.label12.TabIndex = 20;
			this.label12.Text = "Sub Species";
			// 
			// label11
			// 
			this.label11.Location = new System.Drawing.Point(104, 279);
			this.label11.Name = "label11";
			this.label11.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.label11.Size = new System.Drawing.Size(160, 16);
			this.label11.TabIndex = 19;
			this.label11.Text = "Variety";
			// 
			// label10
			// 
			this.label10.Location = new System.Drawing.Point(80, 252);
			this.label10.Name = "label10";
			this.label10.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.label10.Size = new System.Drawing.Size(184, 16);
			this.label10.TabIndex = 18;
			this.label10.Text = "Species";
			// 
			// label9
			// 
			this.label9.Location = new System.Drawing.Point(40, 369);
			this.label9.Name = "label9";
			this.label9.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.label9.Size = new System.Drawing.Size(224, 16);
			this.label9.TabIndex = 17;
			this.label9.Text = "Dry To Green Weight Ratio";
			// 
			// label8
			// 
			this.label8.Location = new System.Drawing.Point(104, 339);
			this.label8.Name = "label8";
			this.label8.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.label8.Size = new System.Drawing.Size(160, 16);
			this.label8.TabIndex = 16;
			this.label8.Text = "Oven Dry Weight";
			// 
			// label7
			// 
			this.label7.Location = new System.Drawing.Point(32, 220);
			this.label7.Name = "label7";
			this.label7.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.label7.Size = new System.Drawing.Size(232, 16);
			this.label7.TabIndex = 15;
			this.label7.Text = "User Defined Tree Species Group";
			// 
			// txtCommon
			// 
			this.txtCommon.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCommon.Location = new System.Drawing.Point(280, 152);
			this.txtCommon.MaxLength = 50;
			this.txtCommon.Name = "txtCommon";
			this.txtCommon.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtCommon.Size = new System.Drawing.Size(296, 23);
			this.txtCommon.TabIndex = 2;
			this.txtCommon.Text = "";
			// 
			// txtSpCd
			// 
			this.txtSpCd.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSpCd.Location = new System.Drawing.Point(280, 120);
			this.txtSpCd.MaxLength = 3;
			this.txtSpCd.Name = "txtSpCd";
			this.txtSpCd.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtSpCd.Size = new System.Drawing.Size(56, 23);
			this.txtSpCd.TabIndex = 1;
			this.txtSpCd.Text = "";
			this.txtSpCd.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSpCd_KeyDown);
			this.txtSpCd.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSpCd_KeyPress);
			// 
			// cmbVariant
			// 
			this.cmbVariant.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVariant.Items.AddRange(new object[] {
															"AK - SouthEast Alaska, Coastal BC",
															"BM - Blue Mountains",
															"CA - Inland CA, Southern Cascades ",
															"CI - Central Idaho",
															"CR - Central Rockies",
															"CS - Central States",
															"EC - Eastside Cascades",
															"EM - Eastern Montana",
															"KT - Kootenai/Kaniksu/Tally Lake",
															"LS - Lake States",
															"NC - Northern California (Klamath Mountains)",
															"NE - Northeast",
															"NI - Northern Idaho (Inland Empire)",
															"PN - Pacific Northwest Coast",
															"SE - Southeast",
															"SN - Southhern",
															"SO - South Central Oregon, Northeast California",
															"TT - Tetons",
															"UT - Utah",
															"WC - Weside Cascades",
															"WS - Westside Sierra Nevada",
															""});
			this.cmbVariant.Location = new System.Drawing.Point(280, 88);
			this.cmbVariant.Name = "cmbVariant";
			this.cmbVariant.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.cmbVariant.Size = new System.Drawing.Size(296, 24);
			this.cmbVariant.TabIndex = 0;
			this.cmbVariant.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVariant_KeyDown);
			this.cmbVariant.SelectedIndexChanged += new System.EventHandler(this.cmbVariant_SelectedIndexChanged);
			// 
			// lblID
			// 
			this.lblID.BackColor = System.Drawing.Color.White;
			this.lblID.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblID.Location = new System.Drawing.Point(280, 56);
			this.lblID.Name = "lblID";
			this.lblID.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.lblID.Size = new System.Drawing.Size(64, 24);
			this.lblID.TabIndex = 13;
			this.lblID.Text = "0";
			// 
			// btnCancel
			// 
			this.btnCancel.Location = new System.Drawing.Point(312, 553);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(88, 48);
			this.btnCancel.TabIndex = 6;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// btnOK
			// 
			this.btnOK.Location = new System.Drawing.Point(224, 553);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(88, 48);
			this.btnOK.TabIndex = 5;
			this.btnOK.Text = "OK";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(104, 92);
			this.label6.Name = "label6";
			this.label6.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.label6.Size = new System.Drawing.Size(160, 16);
			this.label6.TabIndex = 8;
			this.label6.Text = "FVS Variant";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(104, 192);
			this.label5.Name = "label5";
			this.label5.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.label5.Size = new System.Drawing.Size(160, 16);
			this.label5.TabIndex = 7;
			this.label5.Text = "Genus";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(112, 160);
			this.label3.Name = "label3";
			this.label3.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.label3.Size = new System.Drawing.Size(152, 16);
			this.label3.TabIndex = 5;
			this.label3.Text = "Common Name";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(104, 120);
			this.label2.Name = "label2";
			this.label2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.label2.Size = new System.Drawing.Size(160, 16);
			this.label2.TabIndex = 4;
			this.label2.Text = "FIA Tree Species Code";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(232, 64);
			this.label1.Name = "label1";
			this.label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.label1.Size = new System.Drawing.Size(32, 24);
			this.label1.TabIndex = 3;
			this.label1.Text = "ID";
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 19);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.lblTitle.Size = new System.Drawing.Size(618, 24);
			this.lblTitle.TabIndex = 2;
			this.lblTitle.Text = "Processor Tree Species Edit";
			// 
			// uc_processor_tree_spc_edit
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_processor_tree_spc_edit";
			this.Size = new System.Drawing.Size(624, 608);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void txtSpCd_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (Char.IsDigit((char)e.KeyValue))
			{
			}
			else
			{
				if (e.KeyCode == Keys.Back)
				{
				}
				else
				{
					e.Handled=true;	
				}
			}
		}

		private void txtConvert_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (Char.IsDigit((char)e.KeyValue))
			{
			}
			else
			{
				if (e.KeyCode == Keys.Back)
				{
				}
				else
				{
					e.Handled=true;	
				}
			}
		}

		private void cmbVariant_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			e.Handled=true;
		}

		private void txtSpCd_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
			}
			else if (e.KeyChar == '\b')
			{
			}
			else
			{
				e.Handled=true;
			}
			
		}

		private void txtConvert_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
			}
			else if (e.KeyChar == '\b')
			{
			}
			else
			{
				e.Handled=true;
			}
		}
		private void val_data()
		{
		
            this.m_intError = 0;
			if (this.cmbVariant.Text.Trim().Length == 0)
			{
				MessageBox.Show("!!Select A Variant!!","FIA Biosum",
					             System.Windows.Forms.MessageBoxButtons.OK,
					             System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				this.cmbVariant.Focus();
				return;
			}

			if (this.txtSpCd.Text.Trim().Length == 0)
			{
				MessageBox.Show("!!Enter A Tree Species!!","FIA Biosum",
					System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				this.txtSpCd.Focus();
				return;
			}

			if (this.txtConvertToSpCd.Text.Trim().Length == 0) 
			{
			}
			else
			{
				if (Convert.ToInt32(this.txtSpCd.Text) > 299)
				{
					
				}
				else if (Convert.ToInt32(this.txtSpCd.Text) < 300)
				{
					
				}
			}




		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.Cancel;
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			this.val_data();
			if (this.m_intError==0)
				((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.OK;
		}

		private void cmbVariant_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.cmbVariant.Text.Trim().Length > 0)
			{
				this.m_strVariant = this.cmbVariant.Text.Trim().Substring(0,2).ToString();

				if (m_strVariant.Trim().Length > 0)
				{

				    this.cmbFvsSpCd.Items.Clear();
					m_ado.m_strSQL = "SELECT  spcd,fvs_species,common_name " + 
						"FROM " + m_strFvsTreeSpcTable + " " + 
						"WHERE LEN(TRIM(fvs_species)) > 0 AND " + 
						"LEN(TRIM(common_name)) > 0 AND " + 
						"TRIM(fvs_variant) = '" + m_strVariant.Trim() +  "' order by spcd;";
			
					m_ado.SqlQueryReader(m_ado.m_OleDbConnection,m_ado.m_OleDbTransaction,m_ado.m_strSQL);
					if (m_ado.m_OleDbDataReader.HasRows)
					{
						while (m_ado.m_OleDbDataReader.Read())
						{
							this.cmbFvsSpCd.Items.Add(Convert.ToString(m_ado.m_OleDbDataReader["fvs_species"]) + " - " + Convert.ToString(m_ado.m_OleDbDataReader["spcd"]) + " - " + Convert.ToString(m_ado.m_OleDbDataReader["common_name"]));

						}
					}
					m_ado.m_OleDbDataReader.Close();
				}
			}
		}

		private void cmbFvsSpCd_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			string strSpCd="";
			//string strFvsSpCd="";
			string strCommonName="";
			int intFirstDash=0;
			int intSecondDash = 0;
			if (this.cmbFvsSpCd.Text.Trim().Length > 0)
			{
				this.m_strVariant = this.cmbVariant.Text.Trim().Substring(0,2).ToString();

				if (m_strVariant.Trim().Length > 0)
				{
				   intFirstDash = this.cmbFvsSpCd.Text.IndexOf("-",0);
                   intSecondDash = this.cmbFvsSpCd.Text.IndexOf("-", intFirstDash+1);
				   strSpCd = this.cmbFvsSpCd.Text.Substring(intFirstDash + 1, intSecondDash - intFirstDash - 1).Trim();
				   strCommonName = this.cmbFvsSpCd.Text.Substring(intSecondDash + 1, this.cmbFvsSpCd.Text.Trim().Length -  intSecondDash - 1).Trim();
					if (strSpCd.Trim().Length > 0)
					{
                   
					
						m_ado.m_strSQL = "SELECT  spcd,fvs_common_name " + 
							"FROM " + m_strFvsTreeSpcTable + " " + 
							"WHERE TRIM(fvs_variant) = '" + m_strVariant.Trim() +  "' AND " + 
							"      TRIM(common_name) = '" + strCommonName.Trim() + "' AND " + 
							"spcd = " + strSpCd + ";";
			            
						m_ado.SqlQueryReader(m_ado.m_OleDbConnection,m_ado.m_OleDbTransaction,m_ado.m_strSQL);
						if (m_ado.m_OleDbDataReader.HasRows)
						{
							m_ado.m_OleDbDataReader.Read();
							this.txtConvertToSpCd.Text = Convert.ToString(m_ado.m_OleDbDataReader["spcd"]);
                            this.txtFvsCommonName.Text = Convert.ToString(m_ado.m_OleDbDataReader["fvs_common_name"]);
						}
						m_ado.m_OleDbDataReader.Close();
					}
				}

			}
		}
	
		public string strId
		{
			set	{ this.lblID.Text = value; }
			get { return this.lblID.Text.ToString().Trim(); }
		}
		public string strSpCd
		{
			set	{ this.txtSpCd.Text = value; }
			get { return this.txtSpCd.Text.ToString(); }
		}
		
		public string strUserDefinedTreeSpeciesGroup
		{
			set	{ this.lblUserTreeSpCd.Text = value; }
			get { return this.lblUserTreeSpCd.Text.ToString(); }
		}
		public string strCommonName
		{
			set	{ this.txtCommon.Text = value; }
			get { return this.txtCommon.Text.ToString(); }
		}
		public string strVariant
		{
			set	{ this.cmbVariant.Text = value; }
			get { return this.cmbVariant.Text.Trim().Substring(0,2).ToString(); }
		}
		public string strFvsSpeciesCode
		{
			set	{ this.cmbFvsSpCd.Text  = value; }
			get { return this.cmbFvsSpCd.Text.Trim().Substring(0,2).ToString(); }
		}
		public string strTreeSpeciesOvenDryWeight
		{
			set	{ this._txtOvenDryWt.Text  = value; }
			get { return this._txtOvenDryWt.Text.ToString(); }
		}
		public string strDryToGreenWeightRatio
		{
			set	{ this._txtGreenWt.Text  = value; }
			get { return this._txtGreenWt.Text.ToString(); }
		}
		public string strTreeSpeciesGenus
		{
			set	{ this.txtGenus.Text  = value; }
			get { return this.txtGenus.Text.ToString(); }
		}
		public string strTreeSpecies
		{
			set	{ this.txtSpecies.Text  = value; }
			get { return this.txtSpecies.Text.ToString(); }
		}
		public string strTreeSpeciesVariety
		{
			set	{ this.txtVariety.Text  = value; }
			get { return this.txtVariety.Text.ToString(); }
		}
		public string strTreeSpeciesSubSpecies
		{
			set	{ this.txtSubspecies.Text  = value; }
			get { return this.txtSubspecies.Text.ToString(); }
		}
		public string strConvertToSpCd
		{
			set	{ this.txtConvertToSpCd.Text = value; }
			get { return this.txtConvertToSpCd.Text.ToString(); }
		}
		public string strFvsCommonName
		{
			set	{ this.txtFvsCommonName.Text = value; }
			get { return this.txtFvsCommonName.Text.ToString(); }
		}
		



		

	}
}
