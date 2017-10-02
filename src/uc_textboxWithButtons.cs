using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_processor_scenario_escalators.
	/// </summary>
	public class uc_textboxWithButtons : System.Windows.Forms.UserControl
	{
    public int m_intError = 0;
    public string m_strError = "";
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Panel panel1;
        private TextBox textBox1;
        private Button btnOK;
        private Panel panel2;
        
		
		
		
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		
		public uc_textboxWithButtons ()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
            textBox1.GotFocus += delegate { textBox1.Select(0, 0); };

			// TODO: Add any initialization after the InitializeComponent call

		}
        public string TextValue
        {
            set { this.textBox1.Text = value; this.textBox1.SelectionLength = 0; }
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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel2);
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(675, 325);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Location = new System.Drawing.Point(3, 48);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(660, 235);
            this.panel1.TabIndex = 32;
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Control;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox1.Font = new System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(0, 0);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox1.Size = new System.Drawing.Size(660, 235);
            this.textBox1.TabIndex = 1;
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(669, 32);
            this.lblTitle.TabIndex = 30;
            this.lblTitle.Text = "TextBox";
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(568, 3);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 0;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btnOK);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(3, 292);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(669, 30);
            this.panel2.TabIndex = 31;
            // 
            // uc_textboxWithButtons
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_textboxWithButtons";
            this.Size = new System.Drawing.Size(675, 325);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.ParentForm.Close();
        }
		
	}
}
