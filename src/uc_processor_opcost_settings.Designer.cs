namespace FIA_Biosum_Manager
{
    partial class uc_processor_opcost_settings
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_processor_opcost_settings));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.grpOpcost = new System.Windows.Forms.GroupBox();
            this.btnOpcost = new System.Windows.Forms.Button();
            this.txtOpcost = new System.Windows.Forms.TextBox();
            this.grpRdir = new System.Windows.Forms.GroupBox();
            this.btnRdir = new System.Windows.Forms.Button();
            this.txtRdir = new System.Windows.Forms.TextBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnHelp = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.grpOpcost.SuspendLayout();
            this.grpRdir.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.btnOK);
            this.groupBox1.Controls.Add(this.grpOpcost);
            this.groupBox1.Controls.Add(this.grpRdir);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(654, 449);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(399, 286);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(124, 92);
            this.btnCancel.TabIndex = 32;
            this.btnCancel.Text = "CANCEL";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(254, 286);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(124, 92);
            this.btnOK.TabIndex = 31;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // grpOpcost
            // 
            this.grpOpcost.Controls.Add(this.btnOpcost);
            this.grpOpcost.Controls.Add(this.txtOpcost);
            this.grpOpcost.Location = new System.Drawing.Point(28, 168);
            this.grpOpcost.Name = "grpOpcost";
            this.grpOpcost.Size = new System.Drawing.Size(610, 58);
            this.grpOpcost.TabIndex = 30;
            this.grpOpcost.TabStop = false;
            this.grpOpcost.Text = "Directory path of the OPCOST.R file name";
            // 
            // btnOpcost
            // 
            this.btnOpcost.Image = ((System.Drawing.Image)(resources.GetObject("btnOpcost.Image")));
            this.btnOpcost.Location = new System.Drawing.Point(572, 18);
            this.btnOpcost.Name = "btnOpcost";
            this.btnOpcost.Size = new System.Drawing.Size(32, 32);
            this.btnOpcost.TabIndex = 1;
            this.btnOpcost.Click += new System.EventHandler(this.btnOpcost_Click);
            // 
            // txtOpcost
            // 
            this.txtOpcost.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOpcost.Location = new System.Drawing.Point(17, 20);
            this.txtOpcost.Name = "txtOpcost";
            this.txtOpcost.Size = new System.Drawing.Size(549, 26);
            this.txtOpcost.TabIndex = 0;
            this.txtOpcost.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtOpcost_KeyDown);
            this.txtOpcost.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtOpcost_KeyPress);
            // 
            // grpRdir
            // 
            this.grpRdir.Controls.Add(this.btnRdir);
            this.grpRdir.Controls.Add(this.txtRdir);
            this.grpRdir.Location = new System.Drawing.Point(28, 73);
            this.grpRdir.Name = "grpRdir";
            this.grpRdir.Size = new System.Drawing.Size(610, 58);
            this.grpRdir.TabIndex = 29;
            this.grpRdir.TabStop = false;
            this.grpRdir.Text = "Directory path of the (i386) RScript.exe location";
            // 
            // btnRdir
            // 
            this.btnRdir.Image = ((System.Drawing.Image)(resources.GetObject("btnRdir.Image")));
            this.btnRdir.Location = new System.Drawing.Point(572, 18);
            this.btnRdir.Name = "btnRdir";
            this.btnRdir.Size = new System.Drawing.Size(32, 32);
            this.btnRdir.TabIndex = 1;
            this.btnRdir.Click += new System.EventHandler(this.btnRdir_Click);
            // 
            // txtRdir
            // 
            this.txtRdir.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRdir.Location = new System.Drawing.Point(17, 20);
            this.txtRdir.Name = "txtRdir";
            this.txtRdir.Size = new System.Drawing.Size(549, 26);
            this.txtRdir.TabIndex = 0;
            this.txtRdir.Enter += new System.EventHandler(this.txtRdir_Enter);
            this.txtRdir.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRdir_KeyDown);
            this.txtRdir.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRdir_KeyPress);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(648, 24);
            this.lblTitle.TabIndex = 2;
            this.lblTitle.Text = "OPCOST Settings";
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(110, 286);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(124, 92);
            this.btnHelp.TabIndex = 33;
            this.btnHelp.Text = "HELP";
            this.btnHelp.UseVisualStyleBackColor = true;
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // uc_processor_opcost_settings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_processor_opcost_settings";
            this.Size = new System.Drawing.Size(654, 449);
            this.groupBox1.ResumeLayout(false);
            this.grpOpcost.ResumeLayout(false);
            this.grpOpcost.PerformLayout();
            this.grpRdir.ResumeLayout(false);
            this.grpRdir.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.GroupBox grpRdir;
        private System.Windows.Forms.Button btnRdir;
        private System.Windows.Forms.TextBox txtRdir;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.GroupBox grpOpcost;
        private System.Windows.Forms.Button btnOpcost;
        private System.Windows.Forms.TextBox txtOpcost;
        private System.Windows.Forms.Button btnHelp;
    }
}
