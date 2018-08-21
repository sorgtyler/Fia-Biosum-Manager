using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for frmLeft.
	/// </summary>
	public class frmLeft : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnDb;
		private System.Windows.Forms.Button btnCoreAnalysis;
		private System.Windows.Forms.Button btnGIS;
		private System.Windows.Forms.Button btnFVS;
		static int intListTopPosition = 0;
		static int intListHtPosition = 0;
		private System.Windows.Forms.ListBox listbox1;
		private System.Windows.Forms.Button btnfrcs;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmLeft()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			intListTopPosition = Math.Abs(this.btnDb.Top - this.btnDb.Height - 10);
			intListHtPosition = this.Height - (this.btnDb.Height * 5);
			this.listbox1.Top = intListTopPosition;
			this.listbox1.Left = this.Width / 2 - (this.listbox1.Width / 2);
			this.btnDb.Enabled = false;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.btnDb = new System.Windows.Forms.Button();
			this.btnCoreAnalysis = new System.Windows.Forms.Button();
			this.btnGIS = new System.Windows.Forms.Button();
			this.btnFVS = new System.Windows.Forms.Button();
			this.listbox1 = new System.Windows.Forms.ListBox();
			this.btnfrcs = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// btnDb
			// 
			this.btnDb.BackColor = System.Drawing.SystemColors.Control;
			this.btnDb.Dock = System.Windows.Forms.DockStyle.Top;
			this.btnDb.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnDb.Location = new System.Drawing.Point(0, 0);
			this.btnDb.Name = "btnDb";
			this.btnDb.Size = new System.Drawing.Size(128, 23);
			this.btnDb.TabIndex = 0;
			this.btnDb.Text = "Databases";
			this.btnDb.Click += new System.EventHandler(this.btnDb_Click);
			// 
			// btnCoreAnalysis
			// 
			this.btnCoreAnalysis.BackColor = System.Drawing.SystemColors.Control;
			this.btnCoreAnalysis.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.btnCoreAnalysis.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCoreAnalysis.Location = new System.Drawing.Point(0, 308);
			this.btnCoreAnalysis.Name = "btnCoreAnalysis";
			this.btnCoreAnalysis.Size = new System.Drawing.Size(128, 24);
			this.btnCoreAnalysis.TabIndex = 2;
			this.btnCoreAnalysis.Text = "Core Analysis";
			this.btnCoreAnalysis.Click += new System.EventHandler(this.btnCoreAnalysis_Click);
			// 
			// btnGIS
			// 
			this.btnGIS.BackColor = System.Drawing.SystemColors.Control;
			this.btnGIS.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.btnGIS.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnGIS.Location = new System.Drawing.Point(0, 284);
			this.btnGIS.Name = "btnGIS";
			this.btnGIS.Size = new System.Drawing.Size(128, 24);
			this.btnGIS.TabIndex = 3;
			this.btnGIS.Text = "GIS";
			this.btnGIS.Click += new System.EventHandler(this.btnGIS_Click);
			// 
			// btnFVS
			// 
			this.btnFVS.BackColor = System.Drawing.SystemColors.Control;
			this.btnFVS.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.btnFVS.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnFVS.Location = new System.Drawing.Point(0, 260);
			this.btnFVS.Name = "btnFVS";
			this.btnFVS.Size = new System.Drawing.Size(128, 24);
			this.btnFVS.TabIndex = 4;
			this.btnFVS.Text = "FVS";
			this.btnFVS.Click += new System.EventHandler(this.btnFVS_Click);
			// 
			// listbox1
			// 
			this.listbox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.listbox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.listbox1.Location = new System.Drawing.Point(0, 24);
			this.listbox1.Name = "listbox1";
			this.listbox1.Size = new System.Drawing.Size(136, 195);
			this.listbox1.TabIndex = 5;
			this.listbox1.Click += new System.EventHandler(this.listbox1_Click);
			// 
			// btnfrcs
			// 
			this.btnfrcs.BackColor = System.Drawing.SystemColors.Control;
			this.btnfrcs.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.btnfrcs.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnfrcs.Location = new System.Drawing.Point(0, 237);
			this.btnfrcs.Name = "btnfrcs";
			this.btnfrcs.Size = new System.Drawing.Size(128, 23);
			this.btnfrcs.TabIndex = 6;
			this.btnfrcs.Text = "FRCS";
			this.btnfrcs.Click += new System.EventHandler(this.btnfrcs_Click);
			// 
			// frmLeft
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.BackColor = System.Drawing.Color.White;
			this.ClientSize = new System.Drawing.Size(128, 332);
			this.Controls.Add(this.btnfrcs);
			this.Controls.Add(this.listbox1);
			this.Controls.Add(this.btnFVS);
			this.Controls.Add(this.btnGIS);
			this.Controls.Add(this.btnCoreAnalysis);
			this.Controls.Add(this.btnDb);
			this.MaximizeBox = false;
			this.Name = "frmLeft";
			this.Resize += new System.EventHandler(this.frmLeft_Resize);
			this.ResumeLayout(false);

		}
		#endregion

		private void btnCoreAnalysis_Click(object sender, System.EventArgs e)
		{
			if (this.btnCoreAnalysis.Enabled == true) 
			{
				this.btnCoreAnalysis.Dock = DockStyle.Top;
				this.btnCoreAnalysis.Enabled=false;
				
				this.btnDb.Dock = DockStyle.Bottom;
				this.btnDb.Enabled = true;
				this.btnFVS.Dock = DockStyle.Bottom;
				this.btnFVS.Enabled = true;
				this.btnGIS.Dock = DockStyle.Bottom;
				this.btnGIS.Enabled = true;
				this.btnfrcs.Dock = DockStyle.Bottom;
				this.btnfrcs.Enabled = true;
				this.listbox1.Items.Clear();
				this.listbox1.Items.Add("Optimization Analysis");
				this.listbox1.Items.Add("Rules");
				this.listbox1.Top = intListTopPosition;
			}			
			
		}
		private void listbox1_Click(object sender, System.EventArgs e)
		{
			if (this.btnCoreAnalysis.Enabled == false) 
			{

				if (listbox1.Text == "Optimization Analysis") 
				{
					
            
				}
				else if (listbox1.Text == "Rules") 
				{
				}
			}
				
		}
		private void btnDb_Click(object sender, System.EventArgs e)
		{
			if (this.btnDb.Enabled == true) 
			{
				this.btnCoreAnalysis.Dock = DockStyle.Bottom;
				this.btnCoreAnalysis.Enabled = true;
				this.btnDb.Dock = DockStyle.Top;
				this.btnDb.Enabled = false;
				this.btnFVS.Dock = DockStyle.Bottom;
				this.btnFVS.Enabled = true;
				this.btnGIS.Dock = DockStyle.Bottom;
				this.btnGIS.Enabled = true;
				this.btnfrcs.Dock = DockStyle.Bottom;
				this.btnfrcs.Enabled = true;
				this.listbox1.Items.Clear();
				this.listbox1.Top = intListTopPosition;
			}

		}

		private void btnFVS_Click(object sender, System.EventArgs e)
		{
			if (this.btnFVS.Enabled == true) 
			{
				this.btnCoreAnalysis.Dock = DockStyle.Bottom;
				this.btnCoreAnalysis.Enabled = true;
				this.btnDb.Dock = DockStyle.Bottom;
				this.btnDb.Enabled = true;
				this.btnFVS.Dock = DockStyle.Top;
				this.btnFVS.Enabled = false;
				this.btnGIS.Dock = DockStyle.Bottom;
				this.btnGIS.Enabled = true;
				this.btnfrcs.Dock = DockStyle.Bottom;
				this.btnfrcs.Enabled = true;
				this.listbox1.Items.Clear();

				this.listbox1.Items.Add("Create FVS Input");
				this.listbox1.Items.Add("Append FVS Output");
				this.listbox1.Top = intListTopPosition;
			}
			

			
		}

		private void btnGIS_Click(object sender, System.EventArgs e)
		{
			if (this.btnGIS.Enabled == true) 
			{
				this.btnCoreAnalysis.Dock = DockStyle.Bottom;
				this.btnCoreAnalysis.Enabled = true;
				this.btnDb.Dock = DockStyle.Bottom;
				this.btnDb.Enabled = true;
				this.btnFVS.Dock = DockStyle.Bottom;
				this.btnFVS.Enabled = true;
				this.btnGIS.Dock = DockStyle.Top;
				this.btnGIS.Enabled = false;
				this.btnfrcs.Dock = DockStyle.Bottom;
				this.btnfrcs.Enabled = true;
				this.listbox1.Items.Clear();
				this.listbox1.Top = intListTopPosition;
			}


		}

		private void frmLeft_Resize(object sender, System.EventArgs e)
		{
			this.listbox1.Height = this.Height - (this.btnDb.Height * 5) - 50;
			this.listbox1.Top = intListTopPosition;
			this.listbox1.Left = this.Width / 2 - (this.listbox1.Width / 2);
			

		}

		private void btnfrcs_Click(object sender, System.EventArgs e)
		{
			if (this.btnfrcs.Enabled == true) 
			{
				this.btnCoreAnalysis.Dock = DockStyle.Bottom;
				this.btnCoreAnalysis.Enabled = true;
				this.btnDb.Dock = DockStyle.Bottom;
				this.btnDb.Enabled = true;
				this.btnFVS.Dock = DockStyle.Bottom;
				this.btnFVS.Enabled = true;
				this.btnGIS.Dock = DockStyle.Bottom;
				this.btnGIS.Enabled = true;
				this.btnfrcs.Dock = DockStyle.Top;
				this.btnfrcs.Enabled = false;
				this.listbox1.Items.Clear();
				this.listbox1.Top = intListTopPosition;
			}
		}

		
	}
}
