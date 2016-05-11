using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_filter_rows_text_datatype.
	/// </summary>
	public class uc_filter_rows_numeric_datatype : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.RadioButton rdoEqual;
		private System.Windows.Forms.RadioButton rdoNotEqual;
		private System.Windows.Forms.TextBox txtSmall;
		private System.Windows.Forms.Label lblSmallest;
		private System.Windows.Forms.Label lblLargest;
		private System.Windows.Forms.TextBox txtLarge;
		private FIA_Biosum_Manager.txtNumeric m_oTextSmall = new txtNumeric(20,4);
		private FIA_Biosum_Manager.txtNumeric m_oTextLarge = new txtNumeric(20,4);
		private System.Windows.Forms.RadioButton rdoGreaterThan;
		private System.Windows.Forms.RadioButton rdoLessThan;
		private System.Windows.Forms.RadioButton rdoBetween;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_filter_rows_numeric_datatype()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

		    this.Controls.Add(m_oTextSmall);
			this.Controls.Add(m_oTextLarge);
			m_oTextSmall.Location = this.txtSmall.Location;
			m_oTextSmall.Size = this.txtSmall.Size;
			this.txtSmall.Visible=false;
			m_oTextSmall.Show();

			m_oTextLarge.Location = this.txtLarge.Location;
			m_oTextLarge.Size = this.txtLarge.Size;
			this.txtLarge.Visible=false;
			m_oTextLarge.Visible=false;



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
		public string GetFilterType()
		{
			if (this.rdoEqual.Checked) return "EQUAL";
			if (this.rdoNotEqual.Checked) return "NOTEQUAL";
			if (this.rdoGreaterThan.Checked) return "GREATERTHAN";
			if (this.rdoLessThan.Checked) return "LESSTHAN";
			if (this.rdoBetween.Checked) return "BETWEEN";
			return "";
		}
		public string GetSmallestText()
		{
			return this.m_oTextSmall.Text;
		}
		public string GetLargestText()
		{
			return this.m_oTextLarge.Text;
		}
	


		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.btnOK = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.rdoEqual = new System.Windows.Forms.RadioButton();
			this.rdoGreaterThan = new System.Windows.Forms.RadioButton();
			this.rdoNotEqual = new System.Windows.Forms.RadioButton();
			this.rdoLessThan = new System.Windows.Forms.RadioButton();
			this.rdoBetween = new System.Windows.Forms.RadioButton();
			this.txtSmall = new System.Windows.Forms.TextBox();
			this.lblSmallest = new System.Windows.Forms.Label();
			this.lblLargest = new System.Windows.Forms.Label();
			this.txtLarge = new System.Windows.Forms.TextBox();
			this.SuspendLayout();
			// 
			// btnOK
			// 
			this.btnOK.Location = new System.Drawing.Point(32, 168);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(96, 40);
			this.btnOK.TabIndex = 0;
			this.btnOK.Text = "OK";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// btnCancel
			// 
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(160, 168);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(96, 40);
			this.btnCancel.TabIndex = 1;
			this.btnCancel.Text = "CANCEL";
			// 
			// rdoEqual
			// 
			this.rdoEqual.Checked = true;
			this.rdoEqual.Location = new System.Drawing.Point(24, 16);
			this.rdoEqual.Name = "rdoEqual";
			this.rdoEqual.Size = new System.Drawing.Size(64, 24);
			this.rdoEqual.TabIndex = 2;
			this.rdoEqual.TabStop = true;
			this.rdoEqual.Text = "Equal";
			// 
			// rdoGreaterThan
			// 
			this.rdoGreaterThan.Location = new System.Drawing.Point(112, 16);
			this.rdoGreaterThan.Name = "rdoGreaterThan";
			this.rdoGreaterThan.TabIndex = 3;
			this.rdoGreaterThan.Text = "Greater Than";
			// 
			// rdoNotEqual
			// 
			this.rdoNotEqual.Location = new System.Drawing.Point(24, 48);
			this.rdoNotEqual.Name = "rdoNotEqual";
			this.rdoNotEqual.Size = new System.Drawing.Size(72, 24);
			this.rdoNotEqual.TabIndex = 4;
			this.rdoNotEqual.Text = "Not Equal";
			// 
			// rdoLessThan
			// 
			this.rdoLessThan.Location = new System.Drawing.Point(112, 48);
			this.rdoLessThan.Name = "rdoLessThan";
			this.rdoLessThan.TabIndex = 5;
			this.rdoLessThan.Text = "Less Than";
			// 
			// rdoBetween
			// 
			this.rdoBetween.Location = new System.Drawing.Point(240, 16);
			this.rdoBetween.Name = "rdoBetween";
			this.rdoBetween.Size = new System.Drawing.Size(80, 24);
			this.rdoBetween.TabIndex = 6;
			this.rdoBetween.Text = "Between";
			this.rdoBetween.CheckedChanged += new System.EventHandler(this.rdoBetween_CheckedChanged);
			// 
			// txtSmall
			// 
			this.txtSmall.Location = new System.Drawing.Point(80, 96);
			this.txtSmall.Name = "txtSmall";
			this.txtSmall.Size = new System.Drawing.Size(152, 20);
			this.txtSmall.TabIndex = 10;
			this.txtSmall.Text = "";
			// 
			// lblSmallest
			// 
			this.lblSmallest.Location = new System.Drawing.Point(24, 96);
			this.lblSmallest.Name = "lblSmallest";
			this.lblSmallest.Size = new System.Drawing.Size(56, 16);
			this.lblSmallest.TabIndex = 11;
			this.lblSmallest.Text = "Smallest:";
			this.lblSmallest.Visible = false;
			// 
			// lblLargest
			// 
			this.lblLargest.Location = new System.Drawing.Point(24, 128);
			this.lblLargest.Name = "lblLargest";
			this.lblLargest.Size = new System.Drawing.Size(56, 16);
			this.lblLargest.TabIndex = 13;
			this.lblLargest.Text = "Largest:";
			this.lblLargest.Visible = false;
			// 
			// txtLarge
			// 
			this.txtLarge.Location = new System.Drawing.Point(80, 128);
			this.txtLarge.Name = "txtLarge";
			this.txtLarge.Size = new System.Drawing.Size(152, 20);
			this.txtLarge.TabIndex = 12;
			this.txtLarge.Text = "";
			this.txtLarge.Visible = false;
			// 
			// uc_filter_rows_numeric_datatype
			// 
			this.Controls.Add(this.lblLargest);
			this.Controls.Add(this.txtLarge);
			this.Controls.Add(this.lblSmallest);
			this.Controls.Add(this.txtSmall);
			this.Controls.Add(this.rdoBetween);
			this.Controls.Add(this.rdoLessThan);
			this.Controls.Add(this.rdoNotEqual);
			this.Controls.Add(this.rdoGreaterThan);
			this.Controls.Add(this.rdoEqual);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.btnOK);
			this.Name = "uc_filter_rows_numeric_datatype";
			this.Size = new System.Drawing.Size(320, 216);
			this.ResumeLayout(false);

		}
		#endregion

		private void rdoBetween_CheckedChanged(object sender, System.EventArgs e)
		{
			this.lblLargest.Visible = rdoBetween.Checked;
			this.m_oTextLarge.Visible = rdoBetween.Checked;
			this.lblSmallest.Visible=rdoBetween.Checked;
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			if (this.m_oTextSmall.Text.Trim().Length < 2)
			{
				MessageBox.Show("No Value","FIA Biosum");
				this.m_oTextSmall.Focus();
				return;
			}
			if (this.rdoBetween.Checked)
			{
				if (this.m_oTextLarge.Text.Trim().Length < 2)
				{
					MessageBox.Show("No Value","FIA Biosum");
					this.m_oTextLarge.Focus();
					return;
				}
				if (Convert.ToDouble(this.m_oTextSmall.Text) >= Convert.ToDouble(this.m_oTextLarge.Text))
				{
					MessageBox.Show("Smaller value cannot be greater than the larger value","FIA Biosum");
					this.m_oTextSmall.Focus();
					return;
				}
			}
			((frmDialog)this.ParentForm).DialogResult = DialogResult.OK;
			((frmDialog)this.ParentForm).Close();
		}
	}
}
