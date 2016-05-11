namespace FIA_Biosum_Manager
{
    partial class uc_processor_scenario_run
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
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
            this.cmbFilter = new System.Windows.Forms.ComboBox();
            this.pnlFileSizeMonitor = new System.Windows.Forms.Panel();
            this.uc_filesize_monitor3 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.uc_filesize_monitor2 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.uc_filesize_monitor1 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.chkSteepSlope = new System.Windows.Forms.CheckBox();
            this.chkLowSlope = new System.Windows.Forms.CheckBox();
            this.btnRun = new System.Windows.Forms.Button();
            this.lblMsg = new System.Windows.Forms.Label();
            this.lstFvsOutput = new System.Windows.Forms.ListView();
            this.btnUncheckAll = new System.Windows.Forms.Button();
            this.btnChkAll = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.chkFRCS = new System.Windows.Forms.CheckBox();
            this.chkOPCOST = new System.Windows.Forms.CheckBox();
            this.lblHarvestCosts = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.pnlFileSizeMonitor.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(755, 557);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.lblHarvestCosts);
            this.panel1.Controls.Add(this.chkOPCOST);
            this.panel1.Controls.Add(this.chkFRCS);
            this.panel1.Controls.Add(this.cmbFilter);
            this.panel1.Controls.Add(this.pnlFileSizeMonitor);
            this.panel1.Controls.Add(this.chkSteepSlope);
            this.panel1.Controls.Add(this.chkLowSlope);
            this.panel1.Controls.Add(this.btnRun);
            this.panel1.Controls.Add(this.lblMsg);
            this.panel1.Controls.Add(this.lstFvsOutput);
            this.panel1.Controls.Add(this.btnUncheckAll);
            this.panel1.Controls.Add(this.btnChkAll);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 48);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(749, 506);
            this.panel1.TabIndex = 32;
            this.panel1.Resize += new System.EventHandler(this.panel1_Resize);
            // 
            // cmbFilter
            // 
            this.cmbFilter.FormattingEnabled = true;
            this.cmbFilter.Location = new System.Drawing.Point(19, 361);
            this.cmbFilter.Name = "cmbFilter";
            this.cmbFilter.Size = new System.Drawing.Size(61, 21);
            this.cmbFilter.TabIndex = 69;
            // 
            // pnlFileSizeMonitor
            // 
            this.pnlFileSizeMonitor.AutoScroll = true;
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor3);
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor2);
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor1);
            this.pnlFileSizeMonitor.Location = new System.Drawing.Point(19, 392);
            this.pnlFileSizeMonitor.Name = "pnlFileSizeMonitor";
            this.pnlFileSizeMonitor.Size = new System.Drawing.Size(627, 99);
            this.pnlFileSizeMonitor.TabIndex = 68;
            // 
            // uc_filesize_monitor3
            // 
            this.uc_filesize_monitor3.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor3.Information = "";
            this.uc_filesize_monitor3.Location = new System.Drawing.Point(429, 10);
            this.uc_filesize_monitor3.Name = "uc_filesize_monitor3";
            this.uc_filesize_monitor3.Size = new System.Drawing.Size(181, 76);
            this.uc_filesize_monitor3.TabIndex = 2;
            this.uc_filesize_monitor3.Visible = false;
            // 
            // uc_filesize_monitor2
            // 
            this.uc_filesize_monitor2.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor2.Information = "";
            this.uc_filesize_monitor2.Location = new System.Drawing.Point(206, 10);
            this.uc_filesize_monitor2.Name = "uc_filesize_monitor2";
            this.uc_filesize_monitor2.Size = new System.Drawing.Size(181, 76);
            this.uc_filesize_monitor2.TabIndex = 1;
            this.uc_filesize_monitor2.Visible = false;
            // 
            // uc_filesize_monitor1
            // 
            this.uc_filesize_monitor1.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor1.Information = "";
            this.uc_filesize_monitor1.Location = new System.Drawing.Point(5, 10);
            this.uc_filesize_monitor1.Name = "uc_filesize_monitor1";
            this.uc_filesize_monitor1.Size = new System.Drawing.Size(181, 76);
            this.uc_filesize_monitor1.TabIndex = 0;
            this.uc_filesize_monitor1.Visible = false;
            // 
            // chkSteepSlope
            // 
            this.chkSteepSlope.AutoSize = true;
            this.chkSteepSlope.Checked = true;
            this.chkSteepSlope.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSteepSlope.Location = new System.Drawing.Point(156, 7);
            this.chkSteepSlope.Name = "chkSteepSlope";
            this.chkSteepSlope.Size = new System.Drawing.Size(130, 17);
            this.chkSteepSlope.TabIndex = 61;
            this.chkSteepSlope.Text = "Process Steep Slopes";
            this.chkSteepSlope.UseVisualStyleBackColor = true;
            this.chkSteepSlope.CheckedChanged += new System.EventHandler(this.chkSteepSlope_CheckedChanged);
            // 
            // chkLowSlope
            // 
            this.chkLowSlope.AutoSize = true;
            this.chkLowSlope.Checked = true;
            this.chkLowSlope.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkLowSlope.Location = new System.Drawing.Point(19, 7);
            this.chkLowSlope.Name = "chkLowSlope";
            this.chkLowSlope.Size = new System.Drawing.Size(122, 17);
            this.chkLowSlope.TabIndex = 60;
            this.chkLowSlope.Text = "Process Low Slopes";
            this.chkLowSlope.UseVisualStyleBackColor = true;
            this.chkLowSlope.CheckedChanged += new System.EventHandler(this.chkLowSlope_CheckedChanged);
            // 
            // btnRun
            // 
            this.btnRun.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRun.Location = new System.Drawing.Point(317, 354);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(126, 32);
            this.btnRun.TabIndex = 59;
            this.btnRun.Text = "Run";
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // lblMsg
            // 
            this.lblMsg.AutoSize = true;
            this.lblMsg.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMsg.Location = new System.Drawing.Point(16, 327);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(158, 16);
            this.lblMsg.TabIndex = 58;
            this.lblMsg.Text = "Run Status Messages";
            this.lblMsg.Visible = false;
            // 
            // lstFvsOutput
            // 
            this.lstFvsOutput.CheckBoxes = true;
            this.lstFvsOutput.GridLines = true;
            this.lstFvsOutput.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstFvsOutput.HideSelection = false;
            this.lstFvsOutput.Location = new System.Drawing.Point(19, 30);
            this.lstFvsOutput.MultiSelect = false;
            this.lstFvsOutput.Name = "lstFvsOutput";
            this.lstFvsOutput.Size = new System.Drawing.Size(714, 294);
            this.lstFvsOutput.TabIndex = 57;
            this.lstFvsOutput.UseCompatibleStateImageBehavior = false;
            this.lstFvsOutput.View = System.Windows.Forms.View.Details;
            // 
            // btnUncheckAll
            // 
            this.btnUncheckAll.Location = new System.Drawing.Point(156, 354);
            this.btnUncheckAll.Name = "btnUncheckAll";
            this.btnUncheckAll.Size = new System.Drawing.Size(81, 32);
            this.btnUncheckAll.TabIndex = 55;
            this.btnUncheckAll.Text = "Uncheck All";
            this.btnUncheckAll.Click += new System.EventHandler(this.btnUncheckAll_Click);
            // 
            // btnChkAll
            // 
            this.btnChkAll.Location = new System.Drawing.Point(86, 354);
            this.btnChkAll.Name = "btnChkAll";
            this.btnChkAll.Size = new System.Drawing.Size(64, 32);
            this.btnChkAll.TabIndex = 54;
            this.btnChkAll.Text = "Check All";
            this.btnChkAll.Click += new System.EventHandler(this.btnChkAll_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(749, 32);
            this.lblTitle.TabIndex = 31;
            this.lblTitle.Text = "Run Processor Scenario";
            // 
            // chkFRCS
            // 
            this.chkFRCS.AutoSize = true;
            this.chkFRCS.Location = new System.Drawing.Point(603, 8);
            this.chkFRCS.Name = "chkFRCS";
            this.chkFRCS.Size = new System.Drawing.Size(54, 17);
            this.chkFRCS.TabIndex = 70;
            this.chkFRCS.Text = "FRCS";
            this.chkFRCS.UseVisualStyleBackColor = true;
            this.chkFRCS.Click += new System.EventHandler(this.chkFRCS_Click);
            // 
            // chkOPCOST
            // 
            this.chkOPCOST.AutoSize = true;
            this.chkOPCOST.Checked = true;
            this.chkOPCOST.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkOPCOST.Location = new System.Drawing.Point(663, 8);
            this.chkOPCOST.Name = "chkOPCOST";
            this.chkOPCOST.Size = new System.Drawing.Size(70, 17);
            this.chkOPCOST.TabIndex = 71;
            this.chkOPCOST.Text = "OPCOST";
            this.chkOPCOST.UseVisualStyleBackColor = true;
            this.chkOPCOST.Click += new System.EventHandler(this.chkOPCOST_Click);
            // 
            // lblHarvestCosts
            // 
            this.lblHarvestCosts.AutoSize = true;
            this.lblHarvestCosts.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHarvestCosts.ForeColor = System.Drawing.Color.ForestGreen;
            this.lblHarvestCosts.Location = new System.Drawing.Point(466, 9);
            this.lblHarvestCosts.Name = "lblHarvestCosts";
            this.lblHarvestCosts.Size = new System.Drawing.Size(131, 13);
            this.lblHarvestCosts.TabIndex = 72;
            this.lblHarvestCosts.Text = "Harvest Cost Options:";
            // 
            // uc_processor_scenario_run
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_processor_scenario_run";
            this.Size = new System.Drawing.Size(755, 557);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.pnlFileSizeMonitor.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Panel panel1;
        public System.Windows.Forms.Label lblTitle;
        public System.Windows.Forms.ListView lstFvsOutput;
        private System.Windows.Forms.Button btnUncheckAll;
        private System.Windows.Forms.Button btnChkAll;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Label lblMsg;
        private System.Windows.Forms.CheckBox chkSteepSlope;
        private System.Windows.Forms.CheckBox chkLowSlope;
        private System.Windows.Forms.Panel pnlFileSizeMonitor;
        private uc_filesize_monitor uc_filesize_monitor3;
        private uc_filesize_monitor uc_filesize_monitor2;
        private uc_filesize_monitor uc_filesize_monitor1;
        private System.Windows.Forms.ComboBox cmbFilter;
        private System.Windows.Forms.Label lblHarvestCosts;
        private System.Windows.Forms.CheckBox chkOPCOST;
        private System.Windows.Forms.CheckBox chkFRCS;
    }
}
