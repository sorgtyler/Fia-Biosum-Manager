using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_processor_scenario_spc_dbh_group_values.
	/// </summary>
	public class uc_processor_scenario_spc_dbh_group_value : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.TextBox txtSpeciesGroup;
		private System.Windows.Forms.TextBox txtDbhGroup;
		private System.Windows.Forms.TextBox txtCubicFootDollarValue;
		private FIA_Biosum_Manager.ValidateNumericValues m_oValidate = new ValidateNumericValues();
		private string m_strCubicFootDollarValueSave="";
        public bool m_bSave = false;
        private CheckBox chkChips;
        private frmProcessorScenario _oFrmProcessorScenario = null;
		
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_processor_scenario_spc_dbh_group_value()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

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
		public string CubicFootDollarValue
		{
			get {return this.txtCubicFootDollarValue.Text.Trim();}
			set {this.txtCubicFootDollarValue.Text = value;m_strCubicFootDollarValueSave=value;}
		}
		public string DbhGroup
		{
			get {return this.txtDbhGroup.Text.Trim();}
			set {this.txtDbhGroup.Text = value;}
		}
		public string SpeciesGroup
		{
			get {return this.txtSpeciesGroup.Text.Trim();}
			set {this.txtSpeciesGroup.Text = value;}
		}
        public string GetWoodBin()
        {
            if (chkChips.Checked) return "C";
            else return "M";
        }
        public bool EnergyWood
        {
            get { return chkChips.Checked; }
            set { chkChips.Checked = value; }
        }
        public void SetWoodBin(string p_strWoodBin)
        {
            if (p_strWoodBin.Trim() == "C")
            {
                chkChips.Checked = true;
            }
            else
                chkChips.Checked = false;
        }
        public frmProcessorScenario ReferenceProcessorScenarioForm
        {
            get { return _oFrmProcessorScenario; }
            set { _oFrmProcessorScenario = value; }
        }

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.txtSpeciesGroup = new System.Windows.Forms.TextBox();
            this.txtDbhGroup = new System.Windows.Forms.TextBox();
            this.txtCubicFootDollarValue = new System.Windows.Forms.TextBox();
            this.chkChips = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // txtSpeciesGroup
            // 
            this.txtSpeciesGroup.Location = new System.Drawing.Point(8, 6);
            this.txtSpeciesGroup.Name = "txtSpeciesGroup";
            this.txtSpeciesGroup.Size = new System.Drawing.Size(192, 20);
            this.txtSpeciesGroup.TabIndex = 0;
            this.txtSpeciesGroup.Enter += new System.EventHandler(this.txtSpeciesGroup_Enter);
            this.txtSpeciesGroup.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSpeciesGroup_KeyPress);
            // 
            // txtDbhGroup
            // 
            this.txtDbhGroup.Location = new System.Drawing.Point(208, 6);
            this.txtDbhGroup.Name = "txtDbhGroup";
            this.txtDbhGroup.Size = new System.Drawing.Size(88, 20);
            this.txtDbhGroup.TabIndex = 1;
            this.txtDbhGroup.Enter += new System.EventHandler(this.txtDbhGroup_Enter);
            this.txtDbhGroup.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDbhGroup_KeyPress);
            // 
            // txtCubicFootDollarValue
            // 
            this.txtCubicFootDollarValue.Location = new System.Drawing.Point(355, 6);
            this.txtCubicFootDollarValue.Name = "txtCubicFootDollarValue";
            this.txtCubicFootDollarValue.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtCubicFootDollarValue.ShortcutsEnabled = false;
            this.txtCubicFootDollarValue.Size = new System.Drawing.Size(112, 20);
            this.txtCubicFootDollarValue.TabIndex = 2;
            this.txtCubicFootDollarValue.Text = "$0.00";
            this.txtCubicFootDollarValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtCubicFootDollarValue.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCubicFootDollarValue_KeyPress);
            this.txtCubicFootDollarValue.Leave += new System.EventHandler(this.txtValue_Leave);
            // 
            // chkChips
            // 
            this.chkChips.AutoSize = true;
            this.chkChips.Location = new System.Drawing.Point(319, 9);
            this.chkChips.Name = "chkChips";
            this.chkChips.Size = new System.Drawing.Size(15, 14);
            this.chkChips.TabIndex = 3;
            this.chkChips.UseVisualStyleBackColor = true;
            this.chkChips.CheckedChanged += new System.EventHandler(this.chkChips_CheckedChanged);
            // 
            // uc_processor_scenario_spc_dbh_group_value
            // 
            this.Controls.Add(this.chkChips);
            this.Controls.Add(this.txtCubicFootDollarValue);
            this.Controls.Add(this.txtDbhGroup);
            this.Controls.Add(this.txtSpeciesGroup);
            this.Name = "uc_processor_scenario_spc_dbh_group_value";
            this.Size = new System.Drawing.Size(474, 32);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		private void txtSpeciesGroup_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void txtDbhGroup_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void txtValue_Leave(object sender, System.EventArgs e)
		{
			m_oValidate.RoundDecimalLength=2;
			m_oValidate.Money=true;
			m_oValidate.NullsAllowed=false;
			m_oValidate.TestForMaxMin=false;
			m_oValidate.TestForMin=true;
			m_oValidate.MinValue=0;
			m_oValidate.ValidateDecimal(txtCubicFootDollarValue.Text);
            if (m_oValidate.m_intError == 0)
            {
                
                                
                txtCubicFootDollarValue.Text = m_oValidate.ReturnValue;

            }
            else
            {
                this.txtCubicFootDollarValue.Text = this.m_strCubicFootDollarValueSave;
                this.txtCubicFootDollarValue.Focus();

            }


		}
		public void SaveValues()
		{
			this.m_strCubicFootDollarValueSave=this.txtCubicFootDollarValue.Text;
		}

		private void txtSpeciesGroup_Enter(object sender, System.EventArgs e)
		{
			txtCubicFootDollarValue.Focus();
		}

		private void txtDbhGroup_Enter(object sender, System.EventArgs e)
		{
			txtCubicFootDollarValue.Focus();
		}

        private void chkChips_CheckedChanged(object sender, EventArgs e)
        {
            if (this.ReferenceProcessorScenarioForm != null) ReferenceProcessorScenarioForm.m_bSave = true;
            if (chkChips.Checked) this.txtCubicFootDollarValue.Enabled = false;
            else this.txtCubicFootDollarValue.Enabled = true;
        }

        private void txtCubicFootDollarValue_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }

		
	}
	public class uc_processor_scenario_spc_dbh_group_value_collection : System.Collections.CollectionBase
	{
		public uc_processor_scenario_spc_dbh_group_value_collection()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public void Add(uc_processor_scenario_spc_dbh_group_value p_uc)
		{
			// vérify if object is not already in
			if (this.List.Contains(p_uc))
				throw new InvalidOperationException();
 
			// adding it
			this.List.Add(p_uc);
 
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
		public FIA_Biosum_Manager.uc_processor_scenario_spc_dbh_group_value Item(int Index)
		{
			// The appropriate item is retrieved from the List object and
			// explicitly cast to the Widget type, then returned to the 
			// caller.
			return (FIA_Biosum_Manager.uc_processor_scenario_spc_dbh_group_value) List[Index];
		}

	}
}
