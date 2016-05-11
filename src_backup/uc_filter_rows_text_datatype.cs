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
	public class uc_filter_rows_text_datatype : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.RadioButton rdoEqual;
		private System.Windows.Forms.RadioButton rdoStart;
		private System.Windows.Forms.RadioButton rdoNotEqual;
		private System.Windows.Forms.RadioButton rdoNotStart;
		private System.Windows.Forms.RadioButton rdoContain;
		private System.Windows.Forms.RadioButton rdoNotContain;
		private System.Windows.Forms.RadioButton rdoEnd;
		private System.Windows.Forms.RadioButton rdoNotEnd;
		private System.Windows.Forms.TextBox txtSearch;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_filter_rows_text_datatype()
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
		public string GetFilterType()
		{
			if (this.rdoEqual.Checked) return "EQUAL";
			if (this.rdoNotEqual.Checked) return "NOTEQUAL";
			if (this.rdoStart.Checked) return "START";
			if (this.rdoNotStart.Checked) return "NOTSTART";
			if (this.rdoContain.Checked) return "CONTAIN";
			if (this.rdoNotContain.Checked) return "NOTCONTAIN";
			if (this.rdoEnd.Checked) return "END";
			if (this.rdoNotEnd.Checked) return "NOTEND";
			return "";
		}
		public string GetText()
		{
			return this.txtSearch.Text.Trim();
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
			this.rdoStart = new System.Windows.Forms.RadioButton();
			this.rdoNotEqual = new System.Windows.Forms.RadioButton();
			this.rdoNotStart = new System.Windows.Forms.RadioButton();
			this.rdoContain = new System.Windows.Forms.RadioButton();
			this.rdoNotContain = new System.Windows.Forms.RadioButton();
			this.rdoEnd = new System.Windows.Forms.RadioButton();
			this.rdoNotEnd = new System.Windows.Forms.RadioButton();
			this.txtSearch = new System.Windows.Forms.TextBox();
			this.SuspendLayout();
			// 
			// btnOK
			// 
			this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btnOK.Location = new System.Drawing.Point(136, 152);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(96, 40);
			this.btnOK.TabIndex = 0;
			this.btnOK.Text = "OK";
			// 
			// btnCancel
			// 
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(256, 152);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(96, 40);
			this.btnCancel.TabIndex = 1;
			this.btnCancel.Text = "CANCEL";
			// 
			// rdoEqual
			// 
			this.rdoEqual.Checked = true;
			this.rdoEqual.Location = new System.Drawing.Point(8, 16);
			this.rdoEqual.Name = "rdoEqual";
			this.rdoEqual.Size = new System.Drawing.Size(64, 24);
			this.rdoEqual.TabIndex = 2;
			this.rdoEqual.TabStop = true;
			this.rdoEqual.Text = "Equal";
			// 
			// rdoStart
			// 
			this.rdoStart.Location = new System.Drawing.Point(88, 16);
			this.rdoStart.Name = "rdoStart";
			this.rdoStart.Size = new System.Drawing.Size(80, 24);
			this.rdoStart.TabIndex = 3;
			this.rdoStart.Text = "Starts With";
			// 
			// rdoNotEqual
			// 
			this.rdoNotEqual.Location = new System.Drawing.Point(8, 48);
			this.rdoNotEqual.Name = "rdoNotEqual";
			this.rdoNotEqual.Size = new System.Drawing.Size(72, 24);
			this.rdoNotEqual.TabIndex = 4;
			this.rdoNotEqual.Text = "Not Equal";
			// 
			// rdoNotStart
			// 
			this.rdoNotStart.Location = new System.Drawing.Point(88, 48);
			this.rdoNotStart.Name = "rdoNotStart";
			this.rdoNotStart.Size = new System.Drawing.Size(128, 24);
			this.rdoNotStart.TabIndex = 5;
			this.rdoNotStart.Text = "Does Not Start With";
			// 
			// rdoContain
			// 
			this.rdoContain.Location = new System.Drawing.Point(224, 16);
			this.rdoContain.Name = "rdoContain";
			this.rdoContain.Size = new System.Drawing.Size(80, 24);
			this.rdoContain.TabIndex = 6;
			this.rdoContain.Text = "Contains";
			// 
			// rdoNotContain
			// 
			this.rdoNotContain.Location = new System.Drawing.Point(224, 48);
			this.rdoNotContain.Name = "rdoNotContain";
			this.rdoNotContain.Size = new System.Drawing.Size(120, 24);
			this.rdoNotContain.TabIndex = 7;
			this.rdoNotContain.Text = "Does Not Contain";
			// 
			// rdoEnd
			// 
			this.rdoEnd.Location = new System.Drawing.Point(344, 16);
			this.rdoEnd.Name = "rdoEnd";
			this.rdoEnd.Size = new System.Drawing.Size(96, 24);
			this.rdoEnd.TabIndex = 8;
			this.rdoEnd.Text = "Ends With";
			// 
			// rdoNotEnd
			// 
			this.rdoNotEnd.Location = new System.Drawing.Point(344, 48);
			this.rdoNotEnd.Name = "rdoNotEnd";
			this.rdoNotEnd.Size = new System.Drawing.Size(128, 24);
			this.rdoNotEnd.TabIndex = 9;
			this.rdoNotEnd.Text = "Does Not End With";
			// 
			// txtSearch
			// 
			this.txtSearch.Location = new System.Drawing.Point(24, 96);
			this.txtSearch.Name = "txtSearch";
			this.txtSearch.Size = new System.Drawing.Size(432, 20);
			this.txtSearch.TabIndex = 10;
			this.txtSearch.Text = "";
			// 
			// uc_filter_rows_text_datatype
			// 
			this.Controls.Add(this.txtSearch);
			this.Controls.Add(this.rdoNotEnd);
			this.Controls.Add(this.rdoEnd);
			this.Controls.Add(this.rdoNotContain);
			this.Controls.Add(this.rdoContain);
			this.Controls.Add(this.rdoNotStart);
			this.Controls.Add(this.rdoNotEqual);
			this.Controls.Add(this.rdoStart);
			this.Controls.Add(this.rdoEqual);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.btnOK);
			this.Name = "uc_filter_rows_text_datatype";
			this.Size = new System.Drawing.Size(488, 208);
			this.ResumeLayout(false);

		}
		#endregion
	}
}
