using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for frmSettings.
	/// </summary>
	public class frmSettings : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox grpGrids;
		private System.Windows.Forms.Button btnGridFont;
		private System.Windows.Forms.Label lblGridFontName;
		private System.Windows.Forms.Label lblGridFontSize;
		private System.Windows.Forms.Label lblGridFontStyle;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Label lblGridFont;
		private System.Windows.Forms.Label lblGridRowBackgroundColor;
		private System.Windows.Forms.Button btnGridAlternateRowBackground;
		private System.Windows.Forms.Button btnGridRowBackgroundColor;
		private System.Windows.Forms.Button btnGridBackground;
		private System.Windows.Forms.Label lblGridBackgroundColor;
		private System.Windows.Forms.Label lblGridAlternateRowBackgroundColor;
		private System.Windows.Forms.Label lblGridRowForegroundColor;
		private System.Windows.Forms.Button btnGridSelectedRowBackgroundColor;
        private System.Windows.Forms.Label lblGridSelectedRowBackgroundColor;
        private GroupBox grpDebug;
        private Label label1;
        private ComboBox cmbDebug;
        private CheckBox chkDebug;
        private GroupBox grpTableRecordCounts;
        private CheckBox chkScenarioProcessorForm;
        private CheckBox chkFVSOutputForm;
        private CheckBox chkFVSInputForm;
		

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmSettings()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
            //
            //GRIDVIEW
            //
			this.lblGridBackgroundColor.BackColor = frmMain.g_oGridViewBackgroundColor;
			this.lblGridAlternateRowBackgroundColor.BackColor=frmMain.g_oGridViewAlternateRowBackgroundColor;
			this.lblGridRowBackgroundColor.BackColor = frmMain.g_oGridViewRowBackgroundColor;
			this.lblGridRowForegroundColor.BackColor = frmMain.g_oGridViewRowForegroundColor;
			this.lblGridSelectedRowBackgroundColor.BackColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
			if (frmMain.g_oGridViewFont != null)
			{
					this.lblGridFont.Font = frmMain.g_oGridViewFont;
				    this.lblGridFontName.Text = frmMain.g_oGridViewFont.Name;
				    this.lblGridFontSize.Text = frmMain.g_oGridViewFont.Size.ToString().Trim();
				    this.lblGridFontStyle.Text = frmMain.g_oGridViewFont.Style.ToString().Trim();
			}
            //
            //DEBUG
            //
            if (frmMain.g_bDebug) this.chkDebug.Checked = true;
            else this.chkDebug.Checked = false;
            cmbDebug.SelectedIndex = frmMain.g_intDebugLevel-1;
            //
            //SUPPRESS TABLE RECORD COUNTS
            //
            chkFVSInputForm.Checked = frmMain.g_bSuppressFVSInputTableRowCount;
            chkFVSOutputForm.Checked = frmMain.g_bSuppressFVSOutputTableRowCount;
            chkScenarioProcessorForm.Checked = frmMain.g_bSuppressProcessorScenarioTableRowCount;

            


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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmSettings));
            this.grpGrids = new System.Windows.Forms.GroupBox();
            this.lblGridSelectedRowBackgroundColor = new System.Windows.Forms.Label();
            this.btnGridSelectedRowBackgroundColor = new System.Windows.Forms.Button();
            this.lblGridRowForegroundColor = new System.Windows.Forms.Label();
            this.lblGridBackgroundColor = new System.Windows.Forms.Label();
            this.btnGridBackground = new System.Windows.Forms.Button();
            this.lblGridFont = new System.Windows.Forms.Label();
            this.lblGridFontStyle = new System.Windows.Forms.Label();
            this.lblGridFontSize = new System.Windows.Forms.Label();
            this.lblGridFontName = new System.Windows.Forms.Label();
            this.lblGridAlternateRowBackgroundColor = new System.Windows.Forms.Label();
            this.lblGridRowBackgroundColor = new System.Windows.Forms.Label();
            this.btnGridFont = new System.Windows.Forms.Button();
            this.btnGridAlternateRowBackground = new System.Windows.Forms.Button();
            this.btnGridRowBackgroundColor = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.grpDebug = new System.Windows.Forms.GroupBox();
            this.chkDebug = new System.Windows.Forms.CheckBox();
            this.cmbDebug = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.grpTableRecordCounts = new System.Windows.Forms.GroupBox();
            this.chkFVSInputForm = new System.Windows.Forms.CheckBox();
            this.chkFVSOutputForm = new System.Windows.Forms.CheckBox();
            this.chkScenarioProcessorForm = new System.Windows.Forms.CheckBox();
            this.grpGrids.SuspendLayout();
            this.grpDebug.SuspendLayout();
            this.grpTableRecordCounts.SuspendLayout();
            this.SuspendLayout();
            // 
            // grpGrids
            // 
            this.grpGrids.Controls.Add(this.lblGridSelectedRowBackgroundColor);
            this.grpGrids.Controls.Add(this.btnGridSelectedRowBackgroundColor);
            this.grpGrids.Controls.Add(this.lblGridRowForegroundColor);
            this.grpGrids.Controls.Add(this.lblGridBackgroundColor);
            this.grpGrids.Controls.Add(this.btnGridBackground);
            this.grpGrids.Controls.Add(this.lblGridFont);
            this.grpGrids.Controls.Add(this.lblGridFontStyle);
            this.grpGrids.Controls.Add(this.lblGridFontSize);
            this.grpGrids.Controls.Add(this.lblGridFontName);
            this.grpGrids.Controls.Add(this.lblGridAlternateRowBackgroundColor);
            this.grpGrids.Controls.Add(this.lblGridRowBackgroundColor);
            this.grpGrids.Controls.Add(this.btnGridFont);
            this.grpGrids.Controls.Add(this.btnGridAlternateRowBackground);
            this.grpGrids.Controls.Add(this.btnGridRowBackgroundColor);
            this.grpGrids.Location = new System.Drawing.Point(12, 74);
            this.grpGrids.Name = "grpGrids";
            this.grpGrids.Size = new System.Drawing.Size(736, 184);
            this.grpGrids.TabIndex = 0;
            this.grpGrids.TabStop = false;
            this.grpGrids.Text = "Grids";
            // 
            // lblGridSelectedRowBackgroundColor
            // 
            this.lblGridSelectedRowBackgroundColor.BackColor = System.Drawing.Color.Blue;
            this.lblGridSelectedRowBackgroundColor.Location = new System.Drawing.Point(653, 152);
            this.lblGridSelectedRowBackgroundColor.Name = "lblGridSelectedRowBackgroundColor";
            this.lblGridSelectedRowBackgroundColor.Size = new System.Drawing.Size(56, 24);
            this.lblGridSelectedRowBackgroundColor.TabIndex = 13;
            // 
            // btnGridSelectedRowBackgroundColor
            // 
            this.btnGridSelectedRowBackgroundColor.Image = ((System.Drawing.Image)(resources.GetObject("btnGridSelectedRowBackgroundColor.Image")));
            this.btnGridSelectedRowBackgroundColor.Location = new System.Drawing.Point(643, 16);
            this.btnGridSelectedRowBackgroundColor.Name = "btnGridSelectedRowBackgroundColor";
            this.btnGridSelectedRowBackgroundColor.Size = new System.Drawing.Size(80, 120);
            this.btnGridSelectedRowBackgroundColor.TabIndex = 12;
            this.btnGridSelectedRowBackgroundColor.Text = "Selected Row Background";
            this.btnGridSelectedRowBackgroundColor.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnGridSelectedRowBackgroundColor.Click += new System.EventHandler(this.btnGridSelectedRowBackgroundColor_Click);
            // 
            // lblGridRowForegroundColor
            // 
            this.lblGridRowForegroundColor.BackColor = System.Drawing.Color.Black;
            this.lblGridRowForegroundColor.Location = new System.Drawing.Point(24, 152);
            this.lblGridRowForegroundColor.Name = "lblGridRowForegroundColor";
            this.lblGridRowForegroundColor.Size = new System.Drawing.Size(56, 24);
            this.lblGridRowForegroundColor.TabIndex = 11;
            // 
            // lblGridBackgroundColor
            // 
            this.lblGridBackgroundColor.BackColor = System.Drawing.Color.White;
            this.lblGridBackgroundColor.Location = new System.Drawing.Point(376, 152);
            this.lblGridBackgroundColor.Name = "lblGridBackgroundColor";
            this.lblGridBackgroundColor.Size = new System.Drawing.Size(56, 24);
            this.lblGridBackgroundColor.TabIndex = 10;
            // 
            // btnGridBackground
            // 
            this.btnGridBackground.Image = ((System.Drawing.Image)(resources.GetObject("btnGridBackground.Image")));
            this.btnGridBackground.Location = new System.Drawing.Point(360, 16);
            this.btnGridBackground.Name = "btnGridBackground";
            this.btnGridBackground.Size = new System.Drawing.Size(80, 120);
            this.btnGridBackground.TabIndex = 9;
            this.btnGridBackground.Text = "Grid Background";
            this.btnGridBackground.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnGridBackground.Click += new System.EventHandler(this.btnGridBackground_Click);
            // 
            // lblGridFont
            // 
            this.lblGridFont.Location = new System.Drawing.Point(111, 120);
            this.lblGridFont.Name = "lblGridFont";
            this.lblGridFont.Size = new System.Drawing.Size(232, 40);
            this.lblGridFont.TabIndex = 8;
            this.lblGridFont.Text = "Font Example";
            // 
            // lblGridFontStyle
            // 
            this.lblGridFontStyle.Location = new System.Drawing.Point(111, 88);
            this.lblGridFontStyle.Name = "lblGridFontStyle";
            this.lblGridFontStyle.Size = new System.Drawing.Size(160, 16);
            this.lblGridFontStyle.TabIndex = 7;
            this.lblGridFontStyle.Text = "Regular";
            // 
            // lblGridFontSize
            // 
            this.lblGridFontSize.Location = new System.Drawing.Point(111, 64);
            this.lblGridFontSize.Name = "lblGridFontSize";
            this.lblGridFontSize.Size = new System.Drawing.Size(160, 16);
            this.lblGridFontSize.TabIndex = 6;
            this.lblGridFontSize.Text = "8.25";
            // 
            // lblGridFontName
            // 
            this.lblGridFontName.Location = new System.Drawing.Point(111, 40);
            this.lblGridFontName.Name = "lblGridFontName";
            this.lblGridFontName.Size = new System.Drawing.Size(168, 16);
            this.lblGridFontName.TabIndex = 5;
            this.lblGridFontName.Text = "Microsoft Sans Serif";
            // 
            // lblGridAlternateRowBackgroundColor
            // 
            this.lblGridAlternateRowBackgroundColor.BackColor = System.Drawing.Color.LightGreen;
            this.lblGridAlternateRowBackgroundColor.Location = new System.Drawing.Point(560, 152);
            this.lblGridAlternateRowBackgroundColor.Name = "lblGridAlternateRowBackgroundColor";
            this.lblGridAlternateRowBackgroundColor.Size = new System.Drawing.Size(56, 24);
            this.lblGridAlternateRowBackgroundColor.TabIndex = 4;
            // 
            // lblGridRowBackgroundColor
            // 
            this.lblGridRowBackgroundColor.BackColor = System.Drawing.Color.White;
            this.lblGridRowBackgroundColor.Location = new System.Drawing.Point(464, 152);
            this.lblGridRowBackgroundColor.Name = "lblGridRowBackgroundColor";
            this.lblGridRowBackgroundColor.Size = new System.Drawing.Size(56, 24);
            this.lblGridRowBackgroundColor.TabIndex = 3;
            // 
            // btnGridFont
            // 
            this.btnGridFont.Image = ((System.Drawing.Image)(resources.GetObject("btnGridFont.Image")));
            this.btnGridFont.Location = new System.Drawing.Point(16, 16);
            this.btnGridFont.Name = "btnGridFont";
            this.btnGridFont.Size = new System.Drawing.Size(80, 120);
            this.btnGridFont.TabIndex = 2;
            this.btnGridFont.Text = "Font";
            this.btnGridFont.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnGridFont.Click += new System.EventHandler(this.btnGridFont_Click);
            // 
            // btnGridAlternateRowBackground
            // 
            this.btnGridAlternateRowBackground.Image = ((System.Drawing.Image)(resources.GetObject("btnGridAlternateRowBackground.Image")));
            this.btnGridAlternateRowBackground.Location = new System.Drawing.Point(552, 16);
            this.btnGridAlternateRowBackground.Name = "btnGridAlternateRowBackground";
            this.btnGridAlternateRowBackground.Size = new System.Drawing.Size(80, 120);
            this.btnGridAlternateRowBackground.TabIndex = 1;
            this.btnGridAlternateRowBackground.Text = "Alternate Row Background";
            this.btnGridAlternateRowBackground.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnGridAlternateRowBackground.Click += new System.EventHandler(this.btnGridAlternateRowBackground_Click);
            // 
            // btnGridRowBackgroundColor
            // 
            this.btnGridRowBackgroundColor.Image = ((System.Drawing.Image)(resources.GetObject("btnGridRowBackgroundColor.Image")));
            this.btnGridRowBackgroundColor.Location = new System.Drawing.Point(456, 16);
            this.btnGridRowBackgroundColor.Name = "btnGridRowBackgroundColor";
            this.btnGridRowBackgroundColor.Size = new System.Drawing.Size(80, 120);
            this.btnGridRowBackgroundColor.TabIndex = 0;
            this.btnGridRowBackgroundColor.Text = "Row Background";
            this.btnGridRowBackgroundColor.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnGridRowBackgroundColor.Click += new System.EventHandler(this.btnGridRowBackgroundColor_Click);
            // 
            // btnOK
            // 
            this.btnOK.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOK.Location = new System.Drawing.Point(12, 12);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(144, 56);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "OK";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(162, 12);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(144, 56);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // grpDebug
            // 
            this.grpDebug.Controls.Add(this.label1);
            this.grpDebug.Controls.Add(this.cmbDebug);
            this.grpDebug.Controls.Add(this.chkDebug);
            this.grpDebug.Location = new System.Drawing.Point(12, 269);
            this.grpDebug.Name = "grpDebug";
            this.grpDebug.Size = new System.Drawing.Size(317, 46);
            this.grpDebug.TabIndex = 3;
            this.grpDebug.TabStop = false;
            this.grpDebug.Text = "Debug";
            // 
            // chkDebug
            // 
            this.chkDebug.AutoSize = true;
            this.chkDebug.Location = new System.Drawing.Point(16, 21);
            this.chkDebug.Name = "chkDebug";
            this.chkDebug.Size = new System.Drawing.Size(65, 17);
            this.chkDebug.TabIndex = 4;
            this.chkDebug.Text = "Turn On";
            this.chkDebug.UseVisualStyleBackColor = true;
            // 
            // cmbDebug
            // 
            this.cmbDebug.FormattingEnabled = true;
            this.cmbDebug.Items.AddRange(new object[] {
            "1 - Minimal",
            "2 - Some",
            "3 - Maximum"});
            this.cmbDebug.Location = new System.Drawing.Point(125, 19);
            this.cmbDebug.Name = "cmbDebug";
            this.cmbDebug.Size = new System.Drawing.Size(127, 21);
            this.cmbDebug.TabIndex = 4;
            this.cmbDebug.Text = "3 - Maximum";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(86, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(33, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Level";
            // 
            // grpTableRecordCounts
            // 
            this.grpTableRecordCounts.Controls.Add(this.chkScenarioProcessorForm);
            this.grpTableRecordCounts.Controls.Add(this.chkFVSOutputForm);
            this.grpTableRecordCounts.Controls.Add(this.chkFVSInputForm);
            this.grpTableRecordCounts.Location = new System.Drawing.Point(335, 269);
            this.grpTableRecordCounts.Name = "grpTableRecordCounts";
            this.grpTableRecordCounts.Size = new System.Drawing.Size(413, 46);
            this.grpTableRecordCounts.TabIndex = 4;
            this.grpTableRecordCounts.TabStop = false;
            this.grpTableRecordCounts.Text = "Suppress Table Record Counts";
            // 
            // chkFVSInputForm
            // 
            this.chkFVSInputForm.AutoSize = true;
            this.chkFVSInputForm.Location = new System.Drawing.Point(9, 19);
            this.chkFVSInputForm.Name = "chkFVSInputForm";
            this.chkFVSInputForm.Size = new System.Drawing.Size(99, 17);
            this.chkFVSInputForm.TabIndex = 4;
            this.chkFVSInputForm.Text = "FVS Input Form";
            this.chkFVSInputForm.UseVisualStyleBackColor = true;
            // 
            // chkFVSOutputForm
            // 
            this.chkFVSOutputForm.AutoSize = true;
            this.chkFVSOutputForm.Location = new System.Drawing.Point(114, 19);
            this.chkFVSOutputForm.Name = "chkFVSOutputForm";
            this.chkFVSOutputForm.Size = new System.Drawing.Size(107, 17);
            this.chkFVSOutputForm.TabIndex = 5;
            this.chkFVSOutputForm.Text = "FVS Output Form";
            this.chkFVSOutputForm.UseVisualStyleBackColor = true;
            // 
            // chkScenarioProcessorForm
            // 
            this.chkScenarioProcessorForm.AutoSize = true;
            this.chkScenarioProcessorForm.Location = new System.Drawing.Point(227, 19);
            this.chkScenarioProcessorForm.Name = "chkScenarioProcessorForm";
            this.chkScenarioProcessorForm.Size = new System.Drawing.Size(144, 17);
            this.chkScenarioProcessorForm.TabIndex = 6;
            this.chkScenarioProcessorForm.Text = "Processor Scenario Form";
            this.chkScenarioProcessorForm.UseVisualStyleBackColor = true;
            // 
            // frmSettings
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(784, 336);
            this.Controls.Add(this.grpTableRecordCounts);
            this.Controls.Add(this.grpDebug);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.grpGrids);
            this.Name = "frmSettings";
            this.Text = "Settings";
            this.Load += new System.EventHandler(this.frmSettings_Load);
            this.grpGrids.ResumeLayout(false);
            this.grpDebug.ResumeLayout(false);
            this.grpDebug.PerformLayout();
            this.grpTableRecordCounts.ResumeLayout(false);
            this.grpTableRecordCounts.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		private void btnOK_Click(object sender, System.EventArgs e)
		{
            //
            //GRIDVIEW
            //
			frmMain.g_oGridViewAlternateRowBackgroundColor = this.lblGridAlternateRowBackgroundColor.BackColor;
			frmMain.g_oGridViewRowBackgroundColor = this.lblGridRowBackgroundColor.BackColor;
			frmMain.g_oGridViewBackgroundColor=this.lblGridBackgroundColor.BackColor;
			frmMain.g_oGridViewFont = this.lblGridFont.Font;
			frmMain.g_oGridViewRowForegroundColor = this.lblGridRowForegroundColor.BackColor;
			frmMain.g_oGridViewSelectedRowBackgroundColor = this.lblGridSelectedRowBackgroundColor.BackColor;
            //
            //DEBUG
            //
            frmMain.g_bDebug = chkDebug.Checked;
            if (this.cmbDebug.Text.Trim().Length > 0)
            {
                if (this.cmbDebug.Text.Substring(0, 1) == "1") frmMain.g_intDebugLevel = 1;
                if (this.cmbDebug.Text.Substring(0, 1) == "2") frmMain.g_intDebugLevel = 2;
                if (this.cmbDebug.Text.Substring(0, 1) == "3") frmMain.g_intDebugLevel = 3;
            }
            //
            //SUPPRESS TABLE RECORD COUNTS
            //
            frmMain.g_bSuppressFVSInputTableRowCount = chkFVSInputForm.Checked;
            frmMain.g_bSuppressFVSOutputTableRowCount = chkFVSOutputForm.Checked;
            frmMain.g_bSuppressProcessorScenarioTableRowCount = chkScenarioProcessorForm.Checked;
			this.DialogResult=DialogResult.OK;
			this.Close();

		}

		private void btnGridFont_Click(object sender, System.EventArgs e)
		{
			System.Windows.Forms.FontDialog frmTemp = new FontDialog();
			frmTemp.ShowColor=true;
			frmTemp.Font = this.lblGridFont.Font;
			frmTemp.Color = this.lblGridRowForegroundColor.BackColor;
			if(frmTemp.ShowDialog() != DialogResult.Cancel )
			{
				
					this.lblGridFont.Font = frmTemp.Font;
				    this.lblGridFontName.Text = frmTemp.Font.Name;
				    this.lblGridFontSize.Text = frmTemp.Font.Size.ToString().Trim();
				    this.lblGridFontStyle.Text = frmTemp.Font.Style.ToString().Trim();
				    this.lblGridRowForegroundColor.BackColor = frmTemp.Color;
				
			}
			frmTemp.Dispose();
			frmTemp=null;

		}

		private void btnGridAlternateRowBackground_Click(object sender, System.EventArgs e)
		{
			System.Windows.Forms.ColorDialog frmTemp = new ColorDialog();
			if(frmTemp.ShowDialog() != DialogResult.Cancel )
			{
				this.lblGridAlternateRowBackgroundColor.BackColor = frmTemp.Color;
							
				
			}
			frmTemp.Dispose();
			frmTemp=null;

		}

		private void btnGridRowBackgroundColor_Click(object sender, System.EventArgs e)
		{
			System.Windows.Forms.ColorDialog frmTemp = new ColorDialog();
			if(frmTemp.ShowDialog() != DialogResult.Cancel )
			{
				this.lblGridRowBackgroundColor.BackColor = frmTemp.Color;
			}
			frmTemp.Dispose();
			frmTemp=null;


		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			this.DialogResult=DialogResult.Cancel;
			this.Close();
		}

		private void btnGridBackground_Click(object sender, System.EventArgs e)
		{
			System.Windows.Forms.ColorDialog frmTemp = new ColorDialog();
			if(frmTemp.ShowDialog() != DialogResult.Cancel )
			{
				this.lblGridBackgroundColor.BackColor = frmTemp.Color;
							
				
			}
			frmTemp.Dispose();
			frmTemp=null;
		}

		private void btnGridSelectedRowBackgroundColor_Click(object sender, System.EventArgs e)
		{
			System.Windows.Forms.ColorDialog frmTemp = new ColorDialog();
			if(frmTemp.ShowDialog() != DialogResult.Cancel )
			{
				this.lblGridSelectedRowBackgroundColor.BackColor = frmTemp.Color;
			}
			frmTemp.Dispose();
			frmTemp=null;
		
		}

        private void frmSettings_Load(object sender, EventArgs e)
        {

        }

		
	}
}
