namespace FIA_Biosum_Manager
{
    partial class uc_delete_conditions
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_delete_conditions));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.grpboxFilterByState = new System.Windows.Forms.GroupBox();
            this.btnFilterByStateFinish = new System.Windows.Forms.Button();
            this.btnFilterByStateUnselect = new System.Windows.Forms.Button();
            this.btnFilterByStateSelect = new System.Windows.Forms.Button();
            this.lstFilterByState = new System.Windows.Forms.ListView();
            this.btnFilterByStateHelp = new System.Windows.Forms.Button();
            this.btnFilterByStatePrevious = new System.Windows.Forms.Button();
            this.btnFilterByStateNext = new System.Windows.Forms.Button();
            this.btnFilterByStateCancel = new System.Windows.Forms.Button();
            this.grpboxFilterByCondId = new System.Windows.Forms.GroupBox();
            this.btnFilterByPlotFinish = new System.Windows.Forms.Button();
            this.btnFilterByPlotUnselect = new System.Windows.Forms.Button();
            this.btnFilterByPlotSelect = new System.Windows.Forms.Button();
            this.lstFilterByCond = new System.Windows.Forms.ListView();
            this.btnFilterByPlotHelp = new System.Windows.Forms.Button();
            this.btnFilterByPlotPrevious = new System.Windows.Forms.Button();
            this.btnFilterByPlotNext = new System.Windows.Forms.Button();
            this.btnFilterByPlotCancel = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.grpboxFilter = new System.Windows.Forms.GroupBox();
            this.btnFilterFinish = new System.Windows.Forms.Button();
            this.btnFilterHelp = new System.Windows.Forms.Button();
            this.btnFilterPrevious = new System.Windows.Forms.Button();
            this.btnFilterNext = new System.Windows.Forms.Button();
            this.btnFilterCancel = new System.Windows.Forms.Button();
            this.grpboxFilterOptions = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnFilterByFileBrowse = new System.Windows.Forms.Button();
            this.txtFilterByFile = new System.Windows.Forms.TextBox();
            this.rdoFilterByFile = new System.Windows.Forms.RadioButton();
            this.rdoDeleteAllConds = new System.Windows.Forms.RadioButton();
            this.rdoFilterByMenu = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.grpboxFilterByState.SuspendLayout();
            this.grpboxFilterByCondId.SuspendLayout();
            this.grpboxFilter.SuspendLayout();
            this.grpboxFilterOptions.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.grpboxFilterByState);
            this.groupBox1.Controls.Add(this.grpboxFilterByCondId);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Controls.Add(this.grpboxFilter);
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(704, 1159);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // grpboxFilterByState
            // 
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateFinish);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateUnselect);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateSelect);
            this.grpboxFilterByState.Controls.Add(this.lstFilterByState);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateHelp);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStatePrevious);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateNext);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateCancel);
            this.grpboxFilterByState.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.grpboxFilterByState.Location = new System.Drawing.Point(16, 413);
            this.grpboxFilterByState.Name = "grpboxFilterByState";
            this.grpboxFilterByState.Size = new System.Drawing.Size(672, 360);
            this.grpboxFilterByState.TabIndex = 33;
            this.grpboxFilterByState.TabStop = false;
            this.grpboxFilterByState.Text = "Filter By State And County";
            // 
            // btnFilterByStateFinish
            // 
            this.btnFilterByStateFinish.Location = new System.Drawing.Point(584, 325);
            this.btnFilterByStateFinish.Name = "btnFilterByStateFinish";
            this.btnFilterByStateFinish.Size = new System.Drawing.Size(72, 24);
            this.btnFilterByStateFinish.TabIndex = 34;
            this.btnFilterByStateFinish.Text = "Delete";
            this.btnFilterByStateFinish.Click += new System.EventHandler(this.btnFilterFinish_Click);
            // 
            // btnFilterByStateUnselect
            // 
            this.btnFilterByStateUnselect.Location = new System.Drawing.Point(560, 182);
            this.btnFilterByStateUnselect.Name = "btnFilterByStateUnselect";
            this.btnFilterByStateUnselect.Size = new System.Drawing.Size(88, 48);
            this.btnFilterByStateUnselect.TabIndex = 32;
            this.btnFilterByStateUnselect.Text = "Clear All";
            // 
            // btnFilterByStateSelect
            // 
            this.btnFilterByStateSelect.Location = new System.Drawing.Point(560, 118);
            this.btnFilterByStateSelect.Name = "btnFilterByStateSelect";
            this.btnFilterByStateSelect.Size = new System.Drawing.Size(88, 48);
            this.btnFilterByStateSelect.TabIndex = 31;
            this.btnFilterByStateSelect.Text = "Select All";
            // 
            // lstFilterByState
            // 
            this.lstFilterByState.CheckBoxes = true;
            this.lstFilterByState.FullRowSelect = true;
            this.lstFilterByState.GridLines = true;
            this.lstFilterByState.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstFilterByState.HideSelection = false;
            this.lstFilterByState.Location = new System.Drawing.Point(136, 32);
            this.lstFilterByState.Name = "lstFilterByState";
            this.lstFilterByState.Size = new System.Drawing.Size(400, 280);
            this.lstFilterByState.TabIndex = 30;
            this.lstFilterByState.UseCompatibleStateImageBehavior = false;
            this.lstFilterByState.View = System.Windows.Forms.View.Details;
            // 
            // btnFilterByStateHelp
            // 
            this.btnFilterByStateHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnFilterByStateHelp.Location = new System.Drawing.Point(16, 325);
            this.btnFilterByStateHelp.Name = "btnFilterByStateHelp";
            this.btnFilterByStateHelp.Size = new System.Drawing.Size(64, 24);
            this.btnFilterByStateHelp.TabIndex = 23;
            this.btnFilterByStateHelp.Text = "Help";
            // 
            // btnFilterByStatePrevious
            // 
            this.btnFilterByStatePrevious.Location = new System.Drawing.Point(424, 325);
            this.btnFilterByStatePrevious.Name = "btnFilterByStatePrevious";
            this.btnFilterByStatePrevious.Size = new System.Drawing.Size(72, 24);
            this.btnFilterByStatePrevious.TabIndex = 22;
            this.btnFilterByStatePrevious.Text = "< Previous";
            // 
            // btnFilterByStateNext
            // 
            this.btnFilterByStateNext.Location = new System.Drawing.Point(496, 325);
            this.btnFilterByStateNext.Name = "btnFilterByStateNext";
            this.btnFilterByStateNext.Size = new System.Drawing.Size(72, 24);
            this.btnFilterByStateNext.TabIndex = 21;
            this.btnFilterByStateNext.Text = "Next >";
            // 
            // btnFilterByStateCancel
            // 
            this.btnFilterByStateCancel.Location = new System.Drawing.Point(336, 325);
            this.btnFilterByStateCancel.Name = "btnFilterByStateCancel";
            this.btnFilterByStateCancel.Size = new System.Drawing.Size(64, 24);
            this.btnFilterByStateCancel.TabIndex = 20;
            this.btnFilterByStateCancel.Text = "Cancel";
            this.btnFilterByStateCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // grpboxFilterByCondId
            // 
            this.grpboxFilterByCondId.Controls.Add(this.btnFilterByPlotFinish);
            this.grpboxFilterByCondId.Controls.Add(this.btnFilterByPlotUnselect);
            this.grpboxFilterByCondId.Controls.Add(this.btnFilterByPlotSelect);
            this.grpboxFilterByCondId.Controls.Add(this.lstFilterByCond);
            this.grpboxFilterByCondId.Controls.Add(this.btnFilterByPlotHelp);
            this.grpboxFilterByCondId.Controls.Add(this.btnFilterByPlotPrevious);
            this.grpboxFilterByCondId.Controls.Add(this.btnFilterByPlotNext);
            this.grpboxFilterByCondId.Controls.Add(this.btnFilterByPlotCancel);
            this.grpboxFilterByCondId.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.grpboxFilterByCondId.Location = new System.Drawing.Point(16, 781);
            this.grpboxFilterByCondId.Name = "grpboxFilterByCondId";
            this.grpboxFilterByCondId.Size = new System.Drawing.Size(672, 360);
            this.grpboxFilterByCondId.TabIndex = 34;
            this.grpboxFilterByCondId.TabStop = false;
            this.grpboxFilterByCondId.Text = "Filter By Cond";
            // 
            // btnFilterByPlotFinish
            // 
            this.btnFilterByPlotFinish.Location = new System.Drawing.Point(584, 325);
            this.btnFilterByPlotFinish.Name = "btnFilterByPlotFinish";
            this.btnFilterByPlotFinish.Size = new System.Drawing.Size(72, 24);
            this.btnFilterByPlotFinish.TabIndex = 33;
            this.btnFilterByPlotFinish.Text = "Delete";
            this.btnFilterByPlotFinish.Click += new System.EventHandler(this.btnFilterFinish_Click);
            // 
            // btnFilterByPlotUnselect
            // 
            this.btnFilterByPlotUnselect.Location = new System.Drawing.Point(560, 180);
            this.btnFilterByPlotUnselect.Name = "btnFilterByPlotUnselect";
            this.btnFilterByPlotUnselect.Size = new System.Drawing.Size(88, 48);
            this.btnFilterByPlotUnselect.TabIndex = 32;
            this.btnFilterByPlotUnselect.Text = "Clear All";
            // 
            // btnFilterByPlotSelect
            // 
            this.btnFilterByPlotSelect.Location = new System.Drawing.Point(560, 116);
            this.btnFilterByPlotSelect.Name = "btnFilterByPlotSelect";
            this.btnFilterByPlotSelect.Size = new System.Drawing.Size(88, 48);
            this.btnFilterByPlotSelect.TabIndex = 31;
            this.btnFilterByPlotSelect.Text = "Select All";
            // 
            // lstFilterByCond
            // 
            this.lstFilterByCond.CheckBoxes = true;
            this.lstFilterByCond.FullRowSelect = true;
            this.lstFilterByCond.GridLines = true;
            this.lstFilterByCond.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstFilterByCond.HideSelection = false;
            this.lstFilterByCond.Location = new System.Drawing.Point(136, 32);
            this.lstFilterByCond.MultiSelect = false;
            this.lstFilterByCond.Name = "lstFilterByCond";
            this.lstFilterByCond.Size = new System.Drawing.Size(400, 280);
            this.lstFilterByCond.TabIndex = 30;
            this.lstFilterByCond.UseCompatibleStateImageBehavior = false;
            this.lstFilterByCond.View = System.Windows.Forms.View.Details;
            // 
            // btnFilterByPlotHelp
            // 
            this.btnFilterByPlotHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnFilterByPlotHelp.Location = new System.Drawing.Point(16, 325);
            this.btnFilterByPlotHelp.Name = "btnFilterByPlotHelp";
            this.btnFilterByPlotHelp.Size = new System.Drawing.Size(64, 24);
            this.btnFilterByPlotHelp.TabIndex = 23;
            this.btnFilterByPlotHelp.Text = "Help";
            // 
            // btnFilterByPlotPrevious
            // 
            this.btnFilterByPlotPrevious.Location = new System.Drawing.Point(424, 325);
            this.btnFilterByPlotPrevious.Name = "btnFilterByPlotPrevious";
            this.btnFilterByPlotPrevious.Size = new System.Drawing.Size(72, 24);
            this.btnFilterByPlotPrevious.TabIndex = 22;
            this.btnFilterByPlotPrevious.Text = "< Previous";
            // 
            // btnFilterByPlotNext
            // 
            this.btnFilterByPlotNext.Enabled = false;
            this.btnFilterByPlotNext.Location = new System.Drawing.Point(496, 325);
            this.btnFilterByPlotNext.Name = "btnFilterByPlotNext";
            this.btnFilterByPlotNext.Size = new System.Drawing.Size(72, 24);
            this.btnFilterByPlotNext.TabIndex = 21;
            this.btnFilterByPlotNext.Text = "Next >";
            // 
            // btnFilterByPlotCancel
            // 
            this.btnFilterByPlotCancel.Location = new System.Drawing.Point(336, 325);
            this.btnFilterByPlotCancel.Name = "btnFilterByPlotCancel";
            this.btnFilterByPlotCancel.Size = new System.Drawing.Size(64, 24);
            this.btnFilterByPlotCancel.TabIndex = 20;
            this.btnFilterByPlotCancel.Text = "Cancel";
            this.btnFilterByPlotCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(698, 24);
            this.lblTitle.TabIndex = 32;
            this.lblTitle.Text = "Delete Conditions";
            // 
            // grpboxFilter
            // 
            this.grpboxFilter.Controls.Add(this.btnFilterFinish);
            this.grpboxFilter.Controls.Add(this.btnFilterHelp);
            this.grpboxFilter.Controls.Add(this.btnFilterPrevious);
            this.grpboxFilter.Controls.Add(this.btnFilterNext);
            this.grpboxFilter.Controls.Add(this.btnFilterCancel);
            this.grpboxFilter.Controls.Add(this.grpboxFilterOptions);
            this.grpboxFilter.Location = new System.Drawing.Point(16, 46);
            this.grpboxFilter.Name = "grpboxFilter";
            this.grpboxFilter.Size = new System.Drawing.Size(672, 361);
            this.grpboxFilter.TabIndex = 31;
            this.grpboxFilter.TabStop = false;
            this.grpboxFilter.Text = "Filter Options";
            // 
            // btnFilterFinish
            // 
            this.btnFilterFinish.Enabled = false;
            this.btnFilterFinish.Location = new System.Drawing.Point(584, 326);
            this.btnFilterFinish.Name = "btnFilterFinish";
            this.btnFilterFinish.Size = new System.Drawing.Size(72, 24);
            this.btnFilterFinish.TabIndex = 5;
            this.btnFilterFinish.Text = "Delete";
            this.btnFilterFinish.Click += new System.EventHandler(this.btnFilterFinish_Click);
            // 
            // btnFilterHelp
            // 
            this.btnFilterHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnFilterHelp.Location = new System.Drawing.Point(16, 326);
            this.btnFilterHelp.Name = "btnFilterHelp";
            this.btnFilterHelp.Size = new System.Drawing.Size(64, 24);
            this.btnFilterHelp.TabIndex = 2;
            this.btnFilterHelp.Text = "Help";
            this.btnFilterHelp.Visible = false;
            // 
            // btnFilterPrevious
            // 
            this.btnFilterPrevious.Enabled = false;
            this.btnFilterPrevious.Location = new System.Drawing.Point(424, 326);
            this.btnFilterPrevious.Name = "btnFilterPrevious";
            this.btnFilterPrevious.Size = new System.Drawing.Size(72, 24);
            this.btnFilterPrevious.TabIndex = 4;
            this.btnFilterPrevious.Text = "< Previous";
            // 
            // btnFilterNext
            // 
            this.btnFilterNext.Enabled = false;
            this.btnFilterNext.Location = new System.Drawing.Point(496, 326);
            this.btnFilterNext.Name = "btnFilterNext";
            this.btnFilterNext.Size = new System.Drawing.Size(72, 24);
            this.btnFilterNext.TabIndex = 5;
            this.btnFilterNext.Text = "Next >";
            // 
            // btnFilterCancel
            // 
            this.btnFilterCancel.Location = new System.Drawing.Point(336, 326);
            this.btnFilterCancel.Name = "btnFilterCancel";
            this.btnFilterCancel.Size = new System.Drawing.Size(64, 24);
            this.btnFilterCancel.TabIndex = 3;
            this.btnFilterCancel.Text = "Cancel";
            this.btnFilterCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // grpboxFilterOptions
            // 
            this.grpboxFilterOptions.Controls.Add(this.label1);
            this.grpboxFilterOptions.Controls.Add(this.btnFilterByFileBrowse);
            this.grpboxFilterOptions.Controls.Add(this.txtFilterByFile);
            this.grpboxFilterOptions.Controls.Add(this.rdoFilterByFile);
            this.grpboxFilterOptions.Controls.Add(this.rdoFilterByMenu);
            this.grpboxFilterOptions.Controls.Add(this.rdoDeleteAllConds);
            this.grpboxFilterOptions.Location = new System.Drawing.Point(85, 59);
            this.grpboxFilterOptions.Name = "grpboxFilterOptions";
            this.grpboxFilterOptions.Size = new System.Drawing.Size(519, 249);
            this.grpboxFilterOptions.TabIndex = 1;
            this.grpboxFilterOptions.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(6, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(519, 20);
            this.label1.TabIndex = 6;
            this.label1.Text = "Warning: Deleting conditions is irreversible after clicking Delete.";
            // 
            // btnFilterByFileBrowse
            // 
            this.btnFilterByFileBrowse.Enabled = false;
            this.btnFilterByFileBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnFilterByFileBrowse.Image")));
            this.btnFilterByFileBrowse.Location = new System.Drawing.Point(409, 120);
            this.btnFilterByFileBrowse.Name = "btnFilterByFileBrowse";
            this.btnFilterByFileBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnFilterByFileBrowse.TabIndex = 4;
            this.btnFilterByFileBrowse.Click += new System.EventHandler(this.btnFilterByFileBrowse_click);
            // 
            // txtFilterByFile
            // 
            this.txtFilterByFile.Enabled = false;
            this.txtFilterByFile.Location = new System.Drawing.Point(75, 127);
            this.txtFilterByFile.Name = "txtFilterByFile";
            this.txtFilterByFile.Size = new System.Drawing.Size(328, 20);
            this.txtFilterByFile.TabIndex = 3;
            this.txtFilterByFile.TextChanged += new System.EventHandler(this.txtFilterByFile_TextChanged);
            // 
            // rdoFilterByFile
            // 
            this.rdoFilterByFile.Checked = true;
            this.rdoFilterByFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoFilterByFile.Location = new System.Drawing.Point(51, 91);
            this.rdoFilterByFile.Name = "rdoFilterByFile";
            this.rdoFilterByFile.Size = new System.Drawing.Size(432, 32);
            this.rdoFilterByFile.TabIndex = 2;
            this.rdoFilterByFile.TabStop = true;
            this.rdoFilterByFile.Text = "Delete Conditions with Text File of Cond.CN values";
            this.rdoFilterByFile.Click += new System.EventHandler(this.rdoFilterByFile_Click);
            // 
            // rdoDeleteAllConds
            // 
            this.rdoDeleteAllConds.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoDeleteAllConds.Location = new System.Drawing.Point(51, 157);
            this.rdoDeleteAllConds.Name = "rdoDeleteAllConds";
            this.rdoDeleteAllConds.Size = new System.Drawing.Size(400, 32);
            this.rdoDeleteAllConds.TabIndex = 0;
            this.rdoDeleteAllConds.Text = "Delete All Conditions";
            this.rdoDeleteAllConds.Visible = false;
            this.rdoDeleteAllConds.Click += new System.EventHandler(this.rdoDeleteAllConds_Click);
            // 
            // rdoFilterByMenu
            // 
            this.rdoFilterByMenu.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoFilterByMenu.Location = new System.Drawing.Point(51, 182);
            this.rdoFilterByMenu.Name = "rdoFilterByMenu";
            this.rdoFilterByMenu.Size = new System.Drawing.Size(400, 32);
            this.rdoFilterByMenu.TabIndex = 1;
            this.rdoFilterByMenu.Text = "Delete Conditions By Menu Selection";
            this.rdoFilterByMenu.Visible = false;
            this.rdoFilterByMenu.Click += new System.EventHandler(this.rdoFilterByMenu_Click);
            // 
            // uc_delete_conditions
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_delete_conditions";
            this.Size = new System.Drawing.Size(704, 1168);
            this.groupBox1.ResumeLayout(false);
            this.grpboxFilterByState.ResumeLayout(false);
            this.grpboxFilterByCondId.ResumeLayout(false);
            this.grpboxFilter.ResumeLayout(false);
            this.grpboxFilterOptions.ResumeLayout(false);
            this.grpboxFilterOptions.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox grpboxFilter;
        private System.Windows.Forms.Button btnFilterFinish;
        private System.Windows.Forms.Button btnFilterHelp;
        private System.Windows.Forms.Button btnFilterPrevious;
        private System.Windows.Forms.Button btnFilterNext;
        private System.Windows.Forms.Button btnFilterCancel;
        private System.Windows.Forms.GroupBox grpboxFilterOptions;
        private System.Windows.Forms.Button btnFilterByFileBrowse;
        private System.Windows.Forms.TextBox txtFilterByFile;
        private System.Windows.Forms.RadioButton rdoFilterByFile;
        public System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.GroupBox grpboxFilterByState;
        private System.Windows.Forms.Button btnFilterByStateFinish;
        private System.Windows.Forms.Button btnFilterByStateUnselect;
        private System.Windows.Forms.Button btnFilterByStateSelect;
        private System.Windows.Forms.ListView lstFilterByState;
        private System.Windows.Forms.Button btnFilterByStateHelp;
        private System.Windows.Forms.Button btnFilterByStatePrevious;
        private System.Windows.Forms.Button btnFilterByStateNext;
        private System.Windows.Forms.Button btnFilterByStateCancel;
        private System.Windows.Forms.GroupBox grpboxFilterByCondId;
        private System.Windows.Forms.Button btnFilterByPlotFinish;
        private System.Windows.Forms.Button btnFilterByPlotUnselect;
        private System.Windows.Forms.Button btnFilterByPlotSelect;
        private System.Windows.Forms.ListView lstFilterByCond;
        private System.Windows.Forms.Button btnFilterByPlotHelp;
        private System.Windows.Forms.Button btnFilterByPlotPrevious;
        private System.Windows.Forms.Button btnFilterByPlotNext;
        private System.Windows.Forms.Button btnFilterByPlotCancel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton rdoFilterByMenu;
        private System.Windows.Forms.RadioButton rdoDeleteAllConds;
    }
}
