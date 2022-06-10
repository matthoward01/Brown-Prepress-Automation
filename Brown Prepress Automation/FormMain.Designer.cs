namespace Brown_Prepress_Automation
{
    partial class FormMain
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMain));
            this.bSettings = new System.Windows.Forms.Button();
            this.bClearTemp = new System.Windows.Forms.Button();
            this.bStart = new System.Windows.Forms.Button();
            this.bStop = new System.Windows.Forms.Button();
            this.pbMain = new System.Windows.Forms.ProgressBar();
            this.bgwMain = new System.ComponentModel.BackgroundWorker();
            this.tMain = new System.Windows.Forms.Timer(this.components);
            this.rtMain = new System.Windows.Forms.RichTextBox();
            this.pbIndividual = new System.Windows.Forms.ProgressBar();
            this.bgwDownload = new System.ComponentModel.BackgroundWorker();
            this.pbDownload = new System.Windows.Forms.ProgressBar();
            this.lPerc = new System.Windows.Forms.Label();
            this.lDownloaded = new System.Windows.Forms.Label();
            this.lSpeed = new System.Windows.Forms.Label();
            this.lSize = new System.Windows.Forms.Label();
            this.pbSize = new System.Windows.Forms.ProgressBar();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // bSettings
            // 
            this.bSettings.Location = new System.Drawing.Point(12, 12);
            this.bSettings.Name = "bSettings";
            this.bSettings.Size = new System.Drawing.Size(75, 23);
            this.bSettings.TabIndex = 0;
            this.bSettings.Text = "Settings";
            this.bSettings.UseVisualStyleBackColor = true;
            this.bSettings.Click += new System.EventHandler(this.bSettings_Click);
            // 
            // bClearTemp
            // 
            this.bClearTemp.Location = new System.Drawing.Point(436, 12);
            this.bClearTemp.Name = "bClearTemp";
            this.bClearTemp.Size = new System.Drawing.Size(75, 23);
            this.bClearTemp.TabIndex = 1;
            this.bClearTemp.Text = "Clear Temp";
            this.bClearTemp.UseVisualStyleBackColor = true;
            this.bClearTemp.Click += new System.EventHandler(this.bClearTemp_Click);
            // 
            // bStart
            // 
            this.bStart.Location = new System.Drawing.Point(93, 12);
            this.bStart.Name = "bStart";
            this.bStart.Size = new System.Drawing.Size(164, 23);
            this.bStart.TabIndex = 2;
            this.bStart.Text = "Start";
            this.bStart.UseVisualStyleBackColor = true;
            this.bStart.Click += new System.EventHandler(this.bStart_Click);
            // 
            // bStop
            // 
            this.bStop.Location = new System.Drawing.Point(263, 12);
            this.bStop.Name = "bStop";
            this.bStop.Size = new System.Drawing.Size(167, 23);
            this.bStop.TabIndex = 3;
            this.bStop.Text = "Stop";
            this.bStop.UseVisualStyleBackColor = true;
            this.bStop.Visible = false;
            this.bStop.Click += new System.EventHandler(this.bStop_Click);
            // 
            // pbMain
            // 
            this.pbMain.Location = new System.Drawing.Point(12, 274);
            this.pbMain.Name = "pbMain";
            this.pbMain.Size = new System.Drawing.Size(499, 23);
            this.pbMain.TabIndex = 5;
            // 
            // bgwMain
            // 
            this.bgwMain.WorkerReportsProgress = true;
            this.bgwMain.WorkerSupportsCancellation = true;
            this.bgwMain.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgwMain_DoWork);
            this.bgwMain.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bgwMain_ProgressChanged);
            this.bgwMain.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bgwMain_RunWorkerCompleted);
            // 
            // tMain
            // 
            this.tMain.Interval = 10000;
            // 
            // rtMain
            // 
            this.rtMain.BackColor = System.Drawing.SystemColors.Control;
            this.rtMain.ForeColor = System.Drawing.SystemColors.Window;
            this.rtMain.Location = new System.Drawing.Point(12, 41);
            this.rtMain.Name = "rtMain";
            this.rtMain.ReadOnly = true;
            this.rtMain.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
            this.rtMain.Size = new System.Drawing.Size(498, 198);
            this.rtMain.TabIndex = 7;
            this.rtMain.Text = "";
            this.rtMain.TextChanged += new System.EventHandler(this.rtMain_TextChanged);
            // 
            // pbIndividual
            // 
            this.pbIndividual.BackColor = System.Drawing.Color.DimGray;
            this.pbIndividual.Location = new System.Drawing.Point(267, 245);
            this.pbIndividual.Name = "pbIndividual";
            this.pbIndividual.Size = new System.Drawing.Size(244, 23);
            this.pbIndividual.TabIndex = 8;
            // 
            // bgwDownload
            // 
            this.bgwDownload.WorkerReportsProgress = true;
            this.bgwDownload.WorkerSupportsCancellation = true;
            // 
            // pbDownload
            // 
            this.pbDownload.BackColor = System.Drawing.Color.DimGray;
            this.pbDownload.Location = new System.Drawing.Point(12, 245);
            this.pbDownload.Name = "pbDownload";
            this.pbDownload.Size = new System.Drawing.Size(244, 23);
            this.pbDownload.TabIndex = 9;
            // 
            // lPerc
            // 
            this.lPerc.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lPerc.Location = new System.Drawing.Point(201, 216);
            this.lPerc.Name = "lPerc";
            this.lPerc.Size = new System.Drawing.Size(113, 13);
            this.lPerc.TabIndex = 11;
            this.lPerc.Text = "%";
            this.lPerc.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lPerc.Visible = false;
            // 
            // lDownloaded
            // 
            this.lDownloaded.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.lDownloaded.Location = new System.Drawing.Point(106, 160);
            this.lDownloaded.Name = "lDownloaded";
            this.lDownloaded.Size = new System.Drawing.Size(312, 13);
            this.lDownloaded.TabIndex = 10;
            this.lDownloaded.Text = "/";
            this.lDownloaded.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lDownloaded.Visible = false;
            // 
            // lSpeed
            // 
            this.lSpeed.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lSpeed.BackColor = System.Drawing.SystemColors.Control;
            this.lSpeed.Location = new System.Drawing.Point(416, 216);
            this.lSpeed.Name = "lSpeed";
            this.lSpeed.Size = new System.Drawing.Size(94, 13);
            this.lSpeed.TabIndex = 12;
            this.lSpeed.Text = "kb/s";
            this.lSpeed.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            this.lSpeed.Visible = false;
            // 
            // lSize
            // 
            this.lSize.AutoSize = true;
            this.lSize.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lSize.Location = new System.Drawing.Point(436, 12);
            this.lSize.Name = "lSize";
            this.lSize.Size = new System.Drawing.Size(24, 12);
            this.lSize.TabIndex = 13;
            this.lSize.Text = "label";
            this.lSize.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lSize.Visible = false;
            // 
            // pbSize
            // 
            this.pbSize.Location = new System.Drawing.Point(436, 25);
            this.pbSize.Name = "pbSize";
            this.pbSize.Size = new System.Drawing.Size(74, 10);
            this.pbSize.TabIndex = 14;
            this.pbSize.Visible = false;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(435, 245);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 15;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(523, 311);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.pbSize);
            this.Controls.Add(this.lSize);
            this.Controls.Add(this.lSpeed);
            this.Controls.Add(this.lPerc);
            this.Controls.Add(this.lDownloaded);
            this.Controls.Add(this.pbDownload);
            this.Controls.Add(this.pbIndividual);
            this.Controls.Add(this.rtMain);
            this.Controls.Add(this.pbMain);
            this.Controls.Add(this.bStop);
            this.Controls.Add(this.bStart);
            this.Controls.Add(this.bClearTemp);
            this.Controls.Add(this.bSettings);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormMain";
            this.Text = "Brown Prepress Automation (3.3.4.6)";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button bSettings;
        private System.Windows.Forms.Button bClearTemp;
        private System.Windows.Forms.Button bStart;
        private System.Windows.Forms.Button bStop;
        private System.Windows.Forms.ProgressBar pbMain;
        private System.Windows.Forms.Timer tMain;
        public System.Windows.Forms.RichTextBox rtMain;
        public System.ComponentModel.BackgroundWorker bgwMain;
        public System.Windows.Forms.ProgressBar pbIndividual;
        public System.Windows.Forms.ProgressBar pbDownload;
        public System.Windows.Forms.Label lPerc;
        public System.Windows.Forms.Label lDownloaded;
        public System.Windows.Forms.Label lSpeed;
        public System.ComponentModel.BackgroundWorker bgwDownload;
        private System.Windows.Forms.Label lSize;
        private System.Windows.Forms.ProgressBar pbSize;
        private System.Windows.Forms.Button button1;
    }
}

