namespace Brown_Prepress_Automation
{
    partial class FormSettingsShawEmail
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
            this.tbShawEmailList = new System.Windows.Forms.TextBox();
            this.lShawEmailList = new System.Windows.Forms.Label();
            this.lTuftexEmailList = new System.Windows.Forms.Label();
            this.tbTuftexEmailList = new System.Windows.Forms.TextBox();
            this.bEmailSettingSave = new System.Windows.Forms.Button();
            this.bEmailSettingsCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tbShawEmailList
            // 
            this.tbShawEmailList.Location = new System.Drawing.Point(12, 22);
            this.tbShawEmailList.Multiline = true;
            this.tbShawEmailList.Name = "tbShawEmailList";
            this.tbShawEmailList.Size = new System.Drawing.Size(229, 125);
            this.tbShawEmailList.TabIndex = 0;
            // 
            // lShawEmailList
            // 
            this.lShawEmailList.AutoSize = true;
            this.lShawEmailList.Location = new System.Drawing.Point(9, 6);
            this.lShawEmailList.Name = "lShawEmailList";
            this.lShawEmailList.Size = new System.Drawing.Size(152, 13);
            this.lShawEmailList.TabIndex = 7;
            this.lShawEmailList.Text = "Shaw Email List (One Per Line)";
            // 
            // lTuftexEmailList
            // 
            this.lTuftexEmailList.AutoSize = true;
            this.lTuftexEmailList.Location = new System.Drawing.Point(250, 6);
            this.lTuftexEmailList.Name = "lTuftexEmailList";
            this.lTuftexEmailList.Size = new System.Drawing.Size(155, 13);
            this.lTuftexEmailList.TabIndex = 9;
            this.lTuftexEmailList.Text = "Tuftex Email List (One Per Line)";
            // 
            // tbTuftexEmailList
            // 
            this.tbTuftexEmailList.Location = new System.Drawing.Point(253, 22);
            this.tbTuftexEmailList.Multiline = true;
            this.tbTuftexEmailList.Name = "tbTuftexEmailList";
            this.tbTuftexEmailList.Size = new System.Drawing.Size(229, 125);
            this.tbTuftexEmailList.TabIndex = 8;
            // 
            // bEmailSettingSave
            // 
            this.bEmailSettingSave.Location = new System.Drawing.Point(407, 157);
            this.bEmailSettingSave.Name = "bEmailSettingSave";
            this.bEmailSettingSave.Size = new System.Drawing.Size(75, 23);
            this.bEmailSettingSave.TabIndex = 10;
            this.bEmailSettingSave.Text = "Save";
            this.bEmailSettingSave.UseVisualStyleBackColor = true;
            this.bEmailSettingSave.Click += new System.EventHandler(this.bEmailSettingSave_Click);
            // 
            // bEmailSettingsCancel
            // 
            this.bEmailSettingsCancel.Location = new System.Drawing.Point(326, 157);
            this.bEmailSettingsCancel.Name = "bEmailSettingsCancel";
            this.bEmailSettingsCancel.Size = new System.Drawing.Size(75, 23);
            this.bEmailSettingsCancel.TabIndex = 11;
            this.bEmailSettingsCancel.Text = "Cancel";
            this.bEmailSettingsCancel.UseVisualStyleBackColor = true;
            this.bEmailSettingsCancel.Click += new System.EventHandler(this.bEmailSettingsCancel_Click);
            // 
            // FormSettingsEmailAccount
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(494, 195);
            this.Controls.Add(this.bEmailSettingsCancel);
            this.Controls.Add(this.bEmailSettingSave);
            this.Controls.Add(this.lTuftexEmailList);
            this.Controls.Add(this.tbTuftexEmailList);
            this.Controls.Add(this.lShawEmailList);
            this.Controls.Add(this.tbShawEmailList);
            this.Name = "FormSettingsEmailAccount";
            this.Text = "EmailSettings";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbShawEmailList;
        private System.Windows.Forms.Label lShawEmailList;
        private System.Windows.Forms.Label lTuftexEmailList;
        private System.Windows.Forms.TextBox tbTuftexEmailList;
        private System.Windows.Forms.Button bEmailSettingSave;
        private System.Windows.Forms.Button bEmailSettingsCancel;
    }
}