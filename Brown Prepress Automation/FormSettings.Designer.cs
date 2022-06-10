namespace Brown_Prepress_Automation
{
    partial class FormSettings
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
            this.tbHotfolder = new System.Windows.Forms.TextBox();
            this.lHotfolder = new System.Windows.Forms.Label();
            this.bHotfolder = new System.Windows.Forms.Button();
            this.bSettingSave = new System.Windows.Forms.Button();
            this.bSettingCancel = new System.Windows.Forms.Button();
            this.fbdHotfolder = new System.Windows.Forms.FolderBrowserDialog();
            this.bErrorFolder = new System.Windows.Forms.Button();
            this.lErrorFolder = new System.Windows.Forms.Label();
            this.tbErrorFolder = new System.Windows.Forms.TextBox();
            this.fbdErrorFolder = new System.Windows.Forms.FolderBrowserDialog();
            this.bArchiveFolder = new System.Windows.Forms.Button();
            this.lArchiveFolder = new System.Windows.Forms.Label();
            this.tbArchiveFolder = new System.Windows.Forms.TextBox();
            this.fbdArchiveFolder = new System.Windows.Forms.FolderBrowserDialog();
            this.cbAutoUpdate = new System.Windows.Forms.CheckBox();
            this.cbSendEmails = new System.Windows.Forms.CheckBox();
            this.cbDebugOn = new System.Windows.Forms.CheckBox();
            this.bBlueline = new System.Windows.Forms.Button();
            this.lBlueline = new System.Windows.Forms.Label();
            this.tbBlueline = new System.Windows.Forms.TextBox();
            this.bParts = new System.Windows.Forms.Button();
            this.lParts = new System.Windows.Forms.Label();
            this.tbParts = new System.Windows.Forms.TextBox();
            this.fbdHpPaper = new System.Windows.Forms.FolderBrowserDialog();
            this.fbdBlueline = new System.Windows.Forms.FolderBrowserDialog();
            this.fbdParts = new System.Windows.Forms.FolderBrowserDialog();
            this.bTemp = new System.Windows.Forms.Button();
            this.lTemp = new System.Windows.Forms.Label();
            this.tbTemp = new System.Windows.Forms.TextBox();
            this.fbdTemp = new System.Windows.Forms.FolderBrowserDialog();
            this.bEmailSettings = new System.Windows.Forms.Button();
            this.cbPrinters = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.fbd5600Output = new System.Windows.Forms.FolderBrowserDialog();
            this.fbd6800Output = new System.Windows.Forms.FolderBrowserDialog();
            this.bShawTuftex = new System.Windows.Forms.Button();
            this.bResetSettings = new System.Windows.Forms.Button();
            this.fbdHpStyrene = new System.Windows.Forms.FolderBrowserDialog();
            this.bArmstrong = new System.Windows.Forms.Button();
            this.bNourison = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tbHotfolder
            // 
            this.tbHotfolder.Location = new System.Drawing.Point(79, 12);
            this.tbHotfolder.Name = "tbHotfolder";
            this.tbHotfolder.Size = new System.Drawing.Size(506, 20);
            this.tbHotfolder.TabIndex = 0;
            // 
            // lHotfolder
            // 
            this.lHotfolder.AutoSize = true;
            this.lHotfolder.Location = new System.Drawing.Point(12, 15);
            this.lHotfolder.Name = "lHotfolder";
            this.lHotfolder.Size = new System.Drawing.Size(50, 13);
            this.lHotfolder.TabIndex = 1;
            this.lHotfolder.Text = "Hotfolder";
            // 
            // bHotfolder
            // 
            this.bHotfolder.Location = new System.Drawing.Point(591, 12);
            this.bHotfolder.Name = "bHotfolder";
            this.bHotfolder.Size = new System.Drawing.Size(36, 20);
            this.bHotfolder.TabIndex = 2;
            this.bHotfolder.Text = "...";
            this.bHotfolder.UseVisualStyleBackColor = true;
            this.bHotfolder.Click += new System.EventHandler(this.bHotfolder_Click);
            // 
            // bSettingSave
            // 
            this.bSettingSave.Location = new System.Drawing.Point(552, 195);
            this.bSettingSave.Name = "bSettingSave";
            this.bSettingSave.Size = new System.Drawing.Size(75, 23);
            this.bSettingSave.TabIndex = 3;
            this.bSettingSave.Text = "Save";
            this.bSettingSave.UseVisualStyleBackColor = true;
            this.bSettingSave.Click += new System.EventHandler(this.bSettingSave_Click);
            // 
            // bSettingCancel
            // 
            this.bSettingCancel.Location = new System.Drawing.Point(471, 195);
            this.bSettingCancel.Name = "bSettingCancel";
            this.bSettingCancel.Size = new System.Drawing.Size(75, 23);
            this.bSettingCancel.TabIndex = 4;
            this.bSettingCancel.Text = "Cancel";
            this.bSettingCancel.UseVisualStyleBackColor = true;
            this.bSettingCancel.Click += new System.EventHandler(this.bSettingCancel_Click);
            // 
            // fbdHotfolder
            // 
            this.fbdHotfolder.SelectedPath = "\\\\192.168.1.45\\Output1\\BROWN AUTOMATION PROGRAM\\HOTFOLDER\\";
            // 
            // bErrorFolder
            // 
            this.bErrorFolder.Location = new System.Drawing.Point(591, 38);
            this.bErrorFolder.Name = "bErrorFolder";
            this.bErrorFolder.Size = new System.Drawing.Size(36, 20);
            this.bErrorFolder.TabIndex = 7;
            this.bErrorFolder.Text = "...";
            this.bErrorFolder.UseVisualStyleBackColor = true;
            this.bErrorFolder.Click += new System.EventHandler(this.bErrorFolder_Click);
            // 
            // lErrorFolder
            // 
            this.lErrorFolder.AutoSize = true;
            this.lErrorFolder.Location = new System.Drawing.Point(12, 41);
            this.lErrorFolder.Name = "lErrorFolder";
            this.lErrorFolder.Size = new System.Drawing.Size(29, 13);
            this.lErrorFolder.TabIndex = 6;
            this.lErrorFolder.Text = "Error";
            // 
            // tbErrorFolder
            // 
            this.tbErrorFolder.Location = new System.Drawing.Point(79, 38);
            this.tbErrorFolder.Name = "tbErrorFolder";
            this.tbErrorFolder.Size = new System.Drawing.Size(506, 20);
            this.tbErrorFolder.TabIndex = 5;
            // 
            // fbdErrorFolder
            // 
            this.fbdErrorFolder.SelectedPath = "\\\\192.168.1.45\\Output1\\BROWN AUTOMATION PROGRAM\\ERROR\\";
            // 
            // bArchiveFolder
            // 
            this.bArchiveFolder.Location = new System.Drawing.Point(591, 64);
            this.bArchiveFolder.Name = "bArchiveFolder";
            this.bArchiveFolder.Size = new System.Drawing.Size(36, 20);
            this.bArchiveFolder.TabIndex = 10;
            this.bArchiveFolder.Text = "...";
            this.bArchiveFolder.UseVisualStyleBackColor = true;
            this.bArchiveFolder.Click += new System.EventHandler(this.bArchiveFolder_Click);
            // 
            // lArchiveFolder
            // 
            this.lArchiveFolder.AutoSize = true;
            this.lArchiveFolder.Location = new System.Drawing.Point(12, 67);
            this.lArchiveFolder.Name = "lArchiveFolder";
            this.lArchiveFolder.Size = new System.Drawing.Size(43, 13);
            this.lArchiveFolder.TabIndex = 9;
            this.lArchiveFolder.Text = "Archive";
            // 
            // tbArchiveFolder
            // 
            this.tbArchiveFolder.Location = new System.Drawing.Point(79, 64);
            this.tbArchiveFolder.Name = "tbArchiveFolder";
            this.tbArchiveFolder.Size = new System.Drawing.Size(506, 20);
            this.tbArchiveFolder.TabIndex = 8;
            // 
            // fbdArchiveFolder
            // 
            this.fbdArchiveFolder.SelectedPath = "\\\\192.168.1.45\\Output1\\BROWN AUTOMATION PROGRAM\\ARCHIVE\\";
            // 
            // cbAutoUpdate
            // 
            this.cbAutoUpdate.AutoSize = true;
            this.cbAutoUpdate.Location = new System.Drawing.Point(345, 172);
            this.cbAutoUpdate.Name = "cbAutoUpdate";
            this.cbAutoUpdate.Size = new System.Drawing.Size(86, 17);
            this.cbAutoUpdate.TabIndex = 11;
            this.cbAutoUpdate.Text = "Auto Update";
            this.cbAutoUpdate.UseVisualStyleBackColor = true;
            this.cbAutoUpdate.CheckedChanged += new System.EventHandler(this.cbAutoUpdate_CheckedChanged);
            // 
            // cbSendEmails
            // 
            this.cbSendEmails.AutoSize = true;
            this.cbSendEmails.Location = new System.Drawing.Point(431, 172);
            this.cbSendEmails.Name = "cbSendEmails";
            this.cbSendEmails.Size = new System.Drawing.Size(84, 17);
            this.cbSendEmails.TabIndex = 12;
            this.cbSendEmails.Text = "Send Emails";
            this.cbSendEmails.UseVisualStyleBackColor = true;
            this.cbSendEmails.CheckedChanged += new System.EventHandler(this.cbSendEmails_CheckedChanged);
            // 
            // cbDebugOn
            // 
            this.cbDebugOn.AutoSize = true;
            this.cbDebugOn.Location = new System.Drawing.Point(519, 172);
            this.cbDebugOn.Name = "cbDebugOn";
            this.cbDebugOn.Size = new System.Drawing.Size(112, 17);
            this.cbDebugOn.TabIndex = 13;
            this.cbDebugOn.Text = "Debug (Matt Only)";
            this.cbDebugOn.UseVisualStyleBackColor = true;
            this.cbDebugOn.CheckedChanged += new System.EventHandler(this.cbDebugOn_CheckedChanged);
            // 
            // bBlueline
            // 
            this.bBlueline.Location = new System.Drawing.Point(591, 90);
            this.bBlueline.Name = "bBlueline";
            this.bBlueline.Size = new System.Drawing.Size(36, 20);
            this.bBlueline.TabIndex = 19;
            this.bBlueline.Text = "...";
            this.bBlueline.UseVisualStyleBackColor = true;
            this.bBlueline.Click += new System.EventHandler(this.bBlueline_Click);
            // 
            // lBlueline
            // 
            this.lBlueline.AutoSize = true;
            this.lBlueline.Location = new System.Drawing.Point(12, 93);
            this.lBlueline.Name = "lBlueline";
            this.lBlueline.Size = new System.Drawing.Size(44, 13);
            this.lBlueline.TabIndex = 18;
            this.lBlueline.Text = "Blueline";
            // 
            // tbBlueline
            // 
            this.tbBlueline.Location = new System.Drawing.Point(79, 90);
            this.tbBlueline.Name = "tbBlueline";
            this.tbBlueline.Size = new System.Drawing.Size(506, 20);
            this.tbBlueline.TabIndex = 17;
            // 
            // bParts
            // 
            this.bParts.Location = new System.Drawing.Point(591, 116);
            this.bParts.Name = "bParts";
            this.bParts.Size = new System.Drawing.Size(36, 20);
            this.bParts.TabIndex = 22;
            this.bParts.Text = "...";
            this.bParts.UseVisualStyleBackColor = true;
            this.bParts.Click += new System.EventHandler(this.bParts_Click);
            // 
            // lParts
            // 
            this.lParts.AutoSize = true;
            this.lParts.Location = new System.Drawing.Point(12, 119);
            this.lParts.Name = "lParts";
            this.lParts.Size = new System.Drawing.Size(31, 13);
            this.lParts.TabIndex = 21;
            this.lParts.Text = "Parts";
            // 
            // tbParts
            // 
            this.tbParts.Location = new System.Drawing.Point(79, 116);
            this.tbParts.Name = "tbParts";
            this.tbParts.Size = new System.Drawing.Size(506, 20);
            this.tbParts.TabIndex = 20;
            // 
            // fbdHpPaper
            // 
            this.fbdHpPaper.SelectedPath = "\\\\192.168.1.45\\Output1\\CALDERA HOTFOLDERS\\Paper 40x55\\";
            // 
            // fbdBlueline
            // 
            this.fbdBlueline.SelectedPath = "\\\\192.168.1.45\\Output1\\HOTFOLDER - BLUELINE\\";
            // 
            // fbdParts
            // 
            this.fbdParts.SelectedPath = "\\\\192.168.1.45\\Parts\\";
            // 
            // bTemp
            // 
            this.bTemp.Location = new System.Drawing.Point(300, 169);
            this.bTemp.Name = "bTemp";
            this.bTemp.Size = new System.Drawing.Size(34, 20);
            this.bTemp.TabIndex = 28;
            this.bTemp.Text = "...";
            this.bTemp.UseVisualStyleBackColor = true;
            this.bTemp.Click += new System.EventHandler(this.bTemp_Click);
            // 
            // lTemp
            // 
            this.lTemp.AutoSize = true;
            this.lTemp.Location = new System.Drawing.Point(12, 172);
            this.lTemp.Name = "lTemp";
            this.lTemp.Size = new System.Drawing.Size(50, 13);
            this.lTemp.TabIndex = 27;
            this.lTemp.Text = "Temp Dir";
            // 
            // tbTemp
            // 
            this.tbTemp.Location = new System.Drawing.Point(79, 169);
            this.tbTemp.Name = "tbTemp";
            this.tbTemp.Size = new System.Drawing.Size(213, 20);
            this.tbTemp.TabIndex = 26;
            // 
            // fbdTemp
            // 
            this.fbdTemp.SelectedPath = "C:\\Temp\\";
            // 
            // bEmailSettings
            // 
            this.bEmailSettings.Location = new System.Drawing.Point(383, 195);
            this.bEmailSettings.Name = "bEmailSettings";
            this.bEmailSettings.Size = new System.Drawing.Size(82, 23);
            this.bEmailSettings.TabIndex = 29;
            this.bEmailSettings.Text = "Email Settings";
            this.bEmailSettings.UseVisualStyleBackColor = true;
            this.bEmailSettings.Click += new System.EventHandler(this.bEmailSettings_Click);
            // 
            // cbPrinters
            // 
            this.cbPrinters.FormattingEnabled = true;
            this.cbPrinters.Location = new System.Drawing.Point(79, 142);
            this.cbPrinters.Name = "cbPrinters";
            this.cbPrinters.Size = new System.Drawing.Size(213, 21);
            this.cbPrinters.TabIndex = 30;
            this.cbPrinters.SelectionChangeCommitted += new System.EventHandler(this.cbPrinters_SelectionChangeCommitted);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 145);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(37, 13);
            this.label1.TabIndex = 31;
            this.label1.Text = "Printer";
            // 
            // fbd5600Output
            // 
            this.fbd5600Output.SelectedPath = "\\\\192.168.1.45\\Output1\\HOTFOLDER - SHAW AUTO - 5600\\";
            // 
            // fbd6800Output
            // 
            this.fbd6800Output.SelectedPath = "\\\\192.168.1.45\\Output1\\HOTFOLDER - SHAW AUTO - 6800\\";
            // 
            // bShawTuftex
            // 
            this.bShawTuftex.Location = new System.Drawing.Point(99, 195);
            this.bShawTuftex.Name = "bShawTuftex";
            this.bShawTuftex.Size = new System.Drawing.Size(77, 23);
            this.bShawTuftex.TabIndex = 41;
            this.bShawTuftex.Text = "Tuftex/Shaw";
            this.bShawTuftex.UseVisualStyleBackColor = true;
            this.bShawTuftex.Click += new System.EventHandler(this.button1_Click);
            // 
            // bResetSettings
            // 
            this.bResetSettings.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bResetSettings.Location = new System.Drawing.Point(300, 141);
            this.bResetSettings.Name = "bResetSettings";
            this.bResetSettings.Size = new System.Drawing.Size(327, 23);
            this.bResetSettings.TabIndex = 42;
            this.bResetSettings.Text = "Reset All Settings to Default (CAN\'T UNDO)";
            this.bResetSettings.UseVisualStyleBackColor = true;
            this.bResetSettings.Click += new System.EventHandler(this.button2_Click);
            // 
            // fbdHpStyrene
            // 
            this.fbdHpStyrene.SelectedPath = "\\\\192.168.1.45\\Output1\\CALDERA HOTFOLDERS\\Nourison\\";
            // 
            // bArmstrong
            // 
            this.bArmstrong.Location = new System.Drawing.Point(16, 195);
            this.bArmstrong.Name = "bArmstrong";
            this.bArmstrong.Size = new System.Drawing.Size(77, 23);
            this.bArmstrong.TabIndex = 49;
            this.bArmstrong.Text = "Armstrong";
            this.bArmstrong.UseVisualStyleBackColor = true;
            this.bArmstrong.Click += new System.EventHandler(this.bArmstrong_Click);
            // 
            // bNourison
            // 
            this.bNourison.Location = new System.Drawing.Point(182, 195);
            this.bNourison.Name = "bNourison";
            this.bNourison.Size = new System.Drawing.Size(77, 23);
            this.bNourison.TabIndex = 51;
            this.bNourison.Text = "Nourison";
            this.bNourison.UseVisualStyleBackColor = true;
            this.bNourison.Click += new System.EventHandler(this.bNourison_Click);
            // 
            // FormSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(639, 229);
            this.Controls.Add(this.bNourison);
            this.Controls.Add(this.bArmstrong);
            this.Controls.Add(this.bResetSettings);
            this.Controls.Add(this.bShawTuftex);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbPrinters);
            this.Controls.Add(this.bEmailSettings);
            this.Controls.Add(this.bTemp);
            this.Controls.Add(this.lTemp);
            this.Controls.Add(this.tbTemp);
            this.Controls.Add(this.bParts);
            this.Controls.Add(this.lParts);
            this.Controls.Add(this.tbParts);
            this.Controls.Add(this.bBlueline);
            this.Controls.Add(this.lBlueline);
            this.Controls.Add(this.tbBlueline);
            this.Controls.Add(this.cbDebugOn);
            this.Controls.Add(this.cbSendEmails);
            this.Controls.Add(this.cbAutoUpdate);
            this.Controls.Add(this.bArchiveFolder);
            this.Controls.Add(this.lArchiveFolder);
            this.Controls.Add(this.tbArchiveFolder);
            this.Controls.Add(this.bErrorFolder);
            this.Controls.Add(this.lErrorFolder);
            this.Controls.Add(this.tbErrorFolder);
            this.Controls.Add(this.bSettingCancel);
            this.Controls.Add(this.bSettingSave);
            this.Controls.Add(this.bHotfolder);
            this.Controls.Add(this.lHotfolder);
            this.Controls.Add(this.tbHotfolder);
            this.Name = "FormSettings";
            this.Text = "Settings";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbHotfolder;
        private System.Windows.Forms.Label lHotfolder;
        private System.Windows.Forms.Button bHotfolder;
        private System.Windows.Forms.FolderBrowserDialog fbdHotfolder;
        private System.Windows.Forms.Button bErrorFolder;
        private System.Windows.Forms.Label lErrorFolder;
        private System.Windows.Forms.TextBox tbErrorFolder;
        private System.Windows.Forms.FolderBrowserDialog fbdErrorFolder;
        private System.Windows.Forms.Button bArchiveFolder;
        private System.Windows.Forms.Label lArchiveFolder;
        private System.Windows.Forms.TextBox tbArchiveFolder;
        private System.Windows.Forms.FolderBrowserDialog fbdArchiveFolder;
        public System.Windows.Forms.Button bSettingSave;
        public System.Windows.Forms.Button bSettingCancel;
        private System.Windows.Forms.CheckBox cbAutoUpdate;
        private System.Windows.Forms.CheckBox cbSendEmails;
        private System.Windows.Forms.CheckBox cbDebugOn;
        private System.Windows.Forms.Button bBlueline;
        private System.Windows.Forms.Label lBlueline;
        private System.Windows.Forms.TextBox tbBlueline;
        private System.Windows.Forms.Button bParts;
        private System.Windows.Forms.Label lParts;
        private System.Windows.Forms.TextBox tbParts;
        private System.Windows.Forms.FolderBrowserDialog fbdHpPaper;
        private System.Windows.Forms.FolderBrowserDialog fbdBlueline;
        private System.Windows.Forms.FolderBrowserDialog fbdParts;
        private System.Windows.Forms.Button bTemp;
        private System.Windows.Forms.Label lTemp;
        private System.Windows.Forms.TextBox tbTemp;
        private System.Windows.Forms.FolderBrowserDialog fbdTemp;
        private System.Windows.Forms.Button bEmailSettings;
        private System.Windows.Forms.ComboBox cbPrinters;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.FolderBrowserDialog fbd5600Output;
        private System.Windows.Forms.FolderBrowserDialog fbd6800Output;
        private System.Windows.Forms.Button bShawTuftex;
        private System.Windows.Forms.Button bResetSettings;
        private System.Windows.Forms.FolderBrowserDialog fbdHpStyrene;
        private System.Windows.Forms.Button bArmstrong;
        private System.Windows.Forms.Button bNourison;
    }
}