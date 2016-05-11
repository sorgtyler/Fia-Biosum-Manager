namespace FIA_Biosum_Manager
{
    partial class uc_filesize_monitor
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
            this.lblFileName = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.lblMaxSize = new System.Windows.Forms.Label();
            this.lblSize = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.progressBarBasic1 = new ProgressBarBasic.ProgressBarBasic();
            this.btnInfo = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblFileName
            // 
            this.lblFileName.AutoSize = true;
            this.lblFileName.Location = new System.Drawing.Point(61, 19);
            this.lblFileName.Name = "lblFileName";
            this.lblFileName.Size = new System.Drawing.Size(54, 13);
            this.lblFileName.TabIndex = 0;
            this.lblFileName.Text = "File Name";
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(12, 35);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(13, 13);
            this.Label1.TabIndex = 2;
            this.Label1.Text = "0";
            // 
            // lblMaxSize
            // 
            this.lblMaxSize.AutoSize = true;
            this.lblMaxSize.Location = new System.Drawing.Point(145, 35);
            this.lblMaxSize.Name = "lblMaxSize";
            this.lblMaxSize.Size = new System.Drawing.Size(25, 13);
            this.lblMaxSize.TabIndex = 3;
            this.lblMaxSize.Text = "2gb";
            // 
            // lblSize
            // 
            this.lblSize.AutoSize = true;
            this.lblSize.Location = new System.Drawing.Point(61, 53);
            this.lblSize.Name = "lblSize";
            this.lblSize.Size = new System.Drawing.Size(46, 13);
            this.lblSize.TabIndex = 4;
            this.lblSize.Text = "File Size";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.ForestGreen;
            this.label2.Location = new System.Drawing.Point(0, 2);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(101, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "File Size Monitor";
            // 
            // progressBarBasic1
            // 
            this.progressBarBasic1.BackColor = System.Drawing.Color.White;
            this.progressBarBasic1.Font = new System.Drawing.Font("Consolas", 10.25F);
            this.progressBarBasic1.ForeColor = System.Drawing.Color.ForestGreen;
            this.progressBarBasic1.Location = new System.Drawing.Point(31, 35);
            this.progressBarBasic1.Name = "progressBarBasic1";
            this.progressBarBasic1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            this.progressBarBasic1.Size = new System.Drawing.Size(108, 15);
            this.progressBarBasic1.TabIndex = 1;
            this.progressBarBasic1.Text = "progressBarBasic1";
            this.progressBarBasic1.TextStyle = ProgressBarBasic.ProgressBarBasic.TextStyleType.Percentage;
            // 
            // btnInfo
            // 
            this.btnInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnInfo.Location = new System.Drawing.Point(110, 0);
            this.btnInfo.Name = "btnInfo";
            this.btnInfo.Size = new System.Drawing.Size(30, 20);
            this.btnInfo.TabIndex = 6;
            this.btnInfo.Text = "Info";
            this.btnInfo.UseVisualStyleBackColor = true;
            this.btnInfo.Click += new System.EventHandler(this.btnInfo_Click);
            // 
            // uc_filesize_monitor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnInfo);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lblSize);
            this.Controls.Add(this.lblMaxSize);
            this.Controls.Add(this.Label1);
            this.Controls.Add(this.progressBarBasic1);
            this.Controls.Add(this.lblFileName);
            this.ForeColor = System.Drawing.Color.Black;
            this.Name = "uc_filesize_monitor";
            this.Size = new System.Drawing.Size(181, 76);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblFileName;
        private ProgressBarBasic.ProgressBarBasic progressBarBasic1;
        private System.Windows.Forms.Label Label1;
        private System.Windows.Forms.Label label2;
        public System.Windows.Forms.Label lblMaxSize;
        public System.Windows.Forms.Label lblSize;
        private System.Windows.Forms.Button btnInfo;
    }
}
