namespace Brown_Prepress_Automation
{
    partial class FormSettingsNourison
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tbHpOutput = new System.Windows.Forms.TextBox();
            this.lHpOutput = new System.Windows.Forms.Label();
            this.bHpOutput = new System.Windows.Forms.Button();
            this.bSettingSave = new System.Windows.Forms.Button();
            this.bSettingCancel = new System.Windows.Forms.Button();
            this.fbdPdfs = new System.Windows.Forms.FolderBrowserDialog();
            this.fbdHotfolder = new System.Windows.Forms.FolderBrowserDialog();
            this.fbdError = new System.Windows.Forms.FolderBrowserDialog();
            this.fbdArchive = new System.Windows.Forms.FolderBrowserDialog();
            this.lShawHotfolderSettings = new System.Windows.Forms.Label();
            this.bpdfs = new System.Windows.Forms.Button();
            this.lPdfs = new System.Windows.Forms.Label();
            this.tbpdfs = new System.Windows.Forms.TextBox();
            this.bHotfolder = new System.Windows.Forms.Button();
            this.lHotfolder = new System.Windows.Forms.Label();
            this.tbHotfolder = new System.Windows.Forms.TextBox();
            this.bError = new System.Windows.Forms.Button();
            this.lError = new System.Windows.Forms.Label();
            this.tbError = new System.Windows.Forms.TextBox();
            this.fbdHpOutput = new System.Windows.Forms.FolderBrowserDialog();
            this.bArchive = new System.Windows.Forms.Button();
            this.lArchive = new System.Windows.Forms.Label();
            this.tbArchive = new System.Windows.Forms.TextBox();
            this.cbCommon = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // tbHpOutput
            // 
            this.tbHpOutput.Location = new System.Drawing.Point(99, 132);
            this.tbHpOutput.Name = "tbHpOutput";
            this.tbHpOutput.Size = new System.Drawing.Size(487, 20);
            this.tbHpOutput.TabIndex = 0;
            // 
            // lHpOutput
            // 
            this.lHpOutput.AutoSize = true;
            this.lHpOutput.Location = new System.Drawing.Point(14, 135);
            this.lHpOutput.Name = "lHpOutput";
            this.lHpOutput.Size = new System.Drawing.Size(57, 13);
            this.lHpOutput.TabIndex = 1;
            this.lHpOutput.Text = "HP Output";
            // 
            // bHpOutput
            // 
            this.bHpOutput.Location = new System.Drawing.Point(592, 132);
            this.bHpOutput.Name = "bHpOutput";
            this.bHpOutput.Size = new System.Drawing.Size(36, 20);
            this.bHpOutput.TabIndex = 2;
            this.bHpOutput.Text = "...";
            this.bHpOutput.UseVisualStyleBackColor = true;
            this.bHpOutput.Click += new System.EventHandler(this.bHpOutput_Click);
            // 
            // bSettingSave
            // 
            this.bSettingSave.Location = new System.Drawing.Point(552, 158);
            this.bSettingSave.Name = "bSettingSave";
            this.bSettingSave.Size = new System.Drawing.Size(75, 23);
            this.bSettingSave.TabIndex = 3;
            this.bSettingSave.Text = "Save";
            this.bSettingSave.UseVisualStyleBackColor = true;
            this.bSettingSave.Click += new System.EventHandler(this.bSettingSave_Click);
            // 
            // bSettingCancel
            // 
            this.bSettingCancel.Location = new System.Drawing.Point(471, 158);
            this.bSettingCancel.Name = "bSettingCancel";
            this.bSettingCancel.Size = new System.Drawing.Size(75, 23);
            this.bSettingCancel.TabIndex = 4;
            this.bSettingCancel.Text = "Cancel";
            this.bSettingCancel.UseVisualStyleBackColor = true;
            this.bSettingCancel.Click += new System.EventHandler(this.bSettingCancel_Click);
            // 
            // fbdPdfs
            // 
            this.fbdPdfs.SelectedPath = "\\\\192.168.1.45\\Customers\\Nourison\\POP\\Nourison POP\\";
            // 
            // fbdHotfolder
            // 
            this.fbdHotfolder.SelectedPath = "\\\\192.168.1.45\\Output1\\BROWN AUTOMATION PROGRAM\\HOTFOLDER\\Nourison\\";
            // 
            // fbdError
            // 
            this.fbdError.SelectedPath = "\\\\192.168.1.45\\Output1\\BROWN AUTOMATION PROGRAM\\ERROR\\Nourison\\";
            // 
            // fbdArchive
            // 
            this.fbdArchive.SelectedPath = "\\\\192.168.1.45\\Output1\\BROWN AUTOMATION PROGRAM\\Archive\\Nourison\\";
            // 
            // lShawHotfolderSettings
            // 
            this.lShawHotfolderSettings.AutoSize = true;
            this.lShawHotfolderSettings.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lShawHotfolderSettings.Location = new System.Drawing.Point(12, 9);
            this.lShawHotfolderSettings.Name = "lShawHotfolderSettings";
            this.lShawHotfolderSettings.Size = new System.Drawing.Size(109, 13);
            this.lShawHotfolderSettings.TabIndex = 20;
            this.lShawHotfolderSettings.Text = "Hotfolder Settings";
            // 
            // bpdfs
            // 
            this.bpdfs.Location = new System.Drawing.Point(592, 28);
            this.bpdfs.Name = "bpdfs";
            this.bpdfs.Size = new System.Drawing.Size(36, 20);
            this.bpdfs.TabIndex = 23;
            this.bpdfs.Text = "...";
            this.bpdfs.UseVisualStyleBackColor = true;
            this.bpdfs.Click += new System.EventHandler(this.bPdfs_Click);
            // 
            // lPdfs
            // 
            this.lPdfs.AutoSize = true;
            this.lPdfs.Location = new System.Drawing.Point(13, 31);
            this.lPdfs.Name = "lPdfs";
            this.lPdfs.Size = new System.Drawing.Size(73, 13);
            this.lPdfs.TabIndex = 22;
            this.lPdfs.Text = "Nourison Pdfs";
            // 
            // tbpdfs
            // 
            this.tbpdfs.Location = new System.Drawing.Point(99, 28);
            this.tbpdfs.Name = "tbpdfs";
            this.tbpdfs.Size = new System.Drawing.Size(487, 20);
            this.tbpdfs.TabIndex = 21;
            // 
            // bHotfolder
            // 
            this.bHotfolder.Location = new System.Drawing.Point(592, 54);
            this.bHotfolder.Name = "bHotfolder";
            this.bHotfolder.Size = new System.Drawing.Size(36, 20);
            this.bHotfolder.TabIndex = 26;
            this.bHotfolder.Text = "...";
            this.bHotfolder.UseVisualStyleBackColor = true;
            this.bHotfolder.Click += new System.EventHandler(this.bHotfolder_Click);
            // 
            // lHotfolder
            // 
            this.lHotfolder.AutoSize = true;
            this.lHotfolder.Location = new System.Drawing.Point(13, 57);
            this.lHotfolder.Name = "lHotfolder";
            this.lHotfolder.Size = new System.Drawing.Size(50, 13);
            this.lHotfolder.TabIndex = 25;
            this.lHotfolder.Text = "Hotfolder";
            // 
            // tbHotfolder
            // 
            this.tbHotfolder.Location = new System.Drawing.Point(99, 54);
            this.tbHotfolder.Name = "tbHotfolder";
            this.tbHotfolder.Size = new System.Drawing.Size(487, 20);
            this.tbHotfolder.TabIndex = 24;
            // 
            // bError
            // 
            this.bError.Location = new System.Drawing.Point(592, 80);
            this.bError.Name = "bError";
            this.bError.Size = new System.Drawing.Size(36, 20);
            this.bError.TabIndex = 29;
            this.bError.Text = "...";
            this.bError.UseVisualStyleBackColor = true;
            this.bError.Click += new System.EventHandler(this.bError_Click);
            // 
            // lError
            // 
            this.lError.AutoSize = true;
            this.lError.Location = new System.Drawing.Point(13, 83);
            this.lError.Name = "lError";
            this.lError.Size = new System.Drawing.Size(61, 13);
            this.lError.TabIndex = 28;
            this.lError.Text = "Error Folder";
            // 
            // tbError
            // 
            this.tbError.Location = new System.Drawing.Point(99, 80);
            this.tbError.Name = "tbError";
            this.tbError.Size = new System.Drawing.Size(487, 20);
            this.tbError.TabIndex = 27;
            // 
            // fbdHpOutput
            // 
            this.fbdHpOutput.SelectedPath = "\\\\192.168.1.45\\Output1\\CALDERA HOTFOLDERS\\STYRENE - HQ POP\\";
            // 
            // bArchive
            // 
            this.bArchive.Location = new System.Drawing.Point(592, 106);
            this.bArchive.Name = "bArchive";
            this.bArchive.Size = new System.Drawing.Size(36, 20);
            this.bArchive.TabIndex = 32;
            this.bArchive.Text = "...";
            this.bArchive.UseVisualStyleBackColor = true;
            this.bArchive.Click += new System.EventHandler(this.bArchive_Click);
            // 
            // lArchive
            // 
            this.lArchive.AutoSize = true;
            this.lArchive.Location = new System.Drawing.Point(13, 109);
            this.lArchive.Name = "lArchive";
            this.lArchive.Size = new System.Drawing.Size(75, 13);
            this.lArchive.TabIndex = 31;
            this.lArchive.Text = "Archive Folder";
            // 
            // tbArchive
            // 
            this.tbArchive.Location = new System.Drawing.Point(99, 106);
            this.tbArchive.Name = "tbArchive";
            this.tbArchive.Size = new System.Drawing.Size(487, 20);
            this.tbArchive.TabIndex = 30;
            // 
            // cbCommon
            // 
            this.cbCommon.AutoSize = true;
            this.cbCommon.Location = new System.Drawing.Point(15, 164);
            this.cbCommon.Name = "cbCommon";
            this.cbCommon.Size = new System.Drawing.Size(103, 17);
            this.cbCommon.TabIndex = 33;
            this.cbCommon.Text = "Common Boards";
            this.cbCommon.UseVisualStyleBackColor = true;
            // 
            // FormSettingsNourison
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(639, 199);
            this.Controls.Add(this.cbCommon);
            this.Controls.Add(this.bArchive);
            this.Controls.Add(this.lArchive);
            this.Controls.Add(this.tbArchive);
            this.Controls.Add(this.bError);
            this.Controls.Add(this.lError);
            this.Controls.Add(this.tbError);
            this.Controls.Add(this.bHotfolder);
            this.Controls.Add(this.lHotfolder);
            this.Controls.Add(this.tbHotfolder);
            this.Controls.Add(this.bpdfs);
            this.Controls.Add(this.lPdfs);
            this.Controls.Add(this.tbpdfs);
            this.Controls.Add(this.lShawHotfolderSettings);
            this.Controls.Add(this.bSettingCancel);
            this.Controls.Add(this.bSettingSave);
            this.Controls.Add(this.bHpOutput);
            this.Controls.Add(this.lHpOutput);
            this.Controls.Add(this.tbHpOutput);
            this.Name = "FormSettingsNourison";
            this.Text = "Nourison Settings";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbHpOutput;
        private System.Windows.Forms.Label lHpOutput;
        private System.Windows.Forms.Button bHpOutput;
        private System.Windows.Forms.FolderBrowserDialog fbdPdfs;
        private System.Windows.Forms.FolderBrowserDialog fbdHotfolder;
        private System.Windows.Forms.FolderBrowserDialog fbdError;
        public System.Windows.Forms.Button bSettingSave;
        public System.Windows.Forms.Button bSettingCancel;
        private System.Windows.Forms.FolderBrowserDialog fbdArchive;
        private System.Windows.Forms.Label lShawHotfolderSettings;
        private System.Windows.Forms.Button bpdfs;
        private System.Windows.Forms.Label lPdfs;
        private System.Windows.Forms.TextBox tbpdfs;
        private System.Windows.Forms.Button bHotfolder;
        private System.Windows.Forms.Label lHotfolder;
        private System.Windows.Forms.TextBox tbHotfolder;
        private System.Windows.Forms.Button bError;
        private System.Windows.Forms.Label lError;
        private System.Windows.Forms.TextBox tbError;
        private System.Windows.Forms.FolderBrowserDialog fbdHpOutput;
        private System.Windows.Forms.Button bArchive;
        private System.Windows.Forms.Label lArchive;
        private System.Windows.Forms.TextBox tbArchive;
        private System.Windows.Forms.CheckBox cbCommon;
    }
}