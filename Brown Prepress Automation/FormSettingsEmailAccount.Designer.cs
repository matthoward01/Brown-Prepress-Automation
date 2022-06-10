namespace Brown_Prepress_Automation
{
    partial class FormSettingsEmailAccount
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
            this.lExchangeServer = new System.Windows.Forms.Label();
            this.lUsername = new System.Windows.Forms.Label();
            this.lPassword = new System.Windows.Forms.Label();
            this.tbExchangeServer = new System.Windows.Forms.TextBox();
            this.tbUsername = new System.Windows.Forms.TextBox();
            this.tbPassword = new System.Windows.Forms.TextBox();
            this.bEmailSettingSave = new System.Windows.Forms.Button();
            this.bEmailSettingsCancel = new System.Windows.Forms.Button();
            this.lFromEmail = new System.Windows.Forms.Label();
            this.tbFromEmail = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // lExchangeServer
            // 
            this.lExchangeServer.AutoSize = true;
            this.lExchangeServer.Location = new System.Drawing.Point(12, 13);
            this.lExchangeServer.Name = "lExchangeServer";
            this.lExchangeServer.Size = new System.Drawing.Size(89, 13);
            this.lExchangeServer.TabIndex = 1;
            this.lExchangeServer.Text = "Exchange Server";
            // 
            // lUsername
            // 
            this.lUsername.AutoSize = true;
            this.lUsername.Location = new System.Drawing.Point(46, 39);
            this.lUsername.Name = "lUsername";
            this.lUsername.Size = new System.Drawing.Size(55, 13);
            this.lUsername.TabIndex = 2;
            this.lUsername.Text = "Username";
            // 
            // lPassword
            // 
            this.lPassword.AutoSize = true;
            this.lPassword.Location = new System.Drawing.Point(48, 91);
            this.lPassword.Name = "lPassword";
            this.lPassword.Size = new System.Drawing.Size(53, 13);
            this.lPassword.TabIndex = 3;
            this.lPassword.Text = "Password";
            // 
            // tbExchangeServer
            // 
            this.tbExchangeServer.Location = new System.Drawing.Point(107, 10);
            this.tbExchangeServer.Name = "tbExchangeServer";
            this.tbExchangeServer.Size = new System.Drawing.Size(151, 20);
            this.tbExchangeServer.TabIndex = 4;
            this.tbExchangeServer.Text = "bii-owa.brownind.com";
            // 
            // tbUsername
            // 
            this.tbUsername.Location = new System.Drawing.Point(107, 36);
            this.tbUsername.Name = "tbUsername";
            this.tbUsername.Size = new System.Drawing.Size(151, 20);
            this.tbUsername.TabIndex = 5;
            this.tbUsername.Text = "mhoward";
            // 
            // tbPassword
            // 
            this.tbPassword.Location = new System.Drawing.Point(107, 88);
            this.tbPassword.Name = "tbPassword";
            this.tbPassword.PasswordChar = '*';
            this.tbPassword.Size = new System.Drawing.Size(151, 20);
            this.tbPassword.TabIndex = 6;
            this.tbPassword.Text = "mh4094";
            // 
            // bEmailSettingSave
            // 
            this.bEmailSettingSave.Location = new System.Drawing.Point(183, 114);
            this.bEmailSettingSave.Name = "bEmailSettingSave";
            this.bEmailSettingSave.Size = new System.Drawing.Size(75, 23);
            this.bEmailSettingSave.TabIndex = 10;
            this.bEmailSettingSave.Text = "Save";
            this.bEmailSettingSave.UseVisualStyleBackColor = true;
            this.bEmailSettingSave.Click += new System.EventHandler(this.bEmailSettingSave_Click);
            // 
            // bEmailSettingsCancel
            // 
            this.bEmailSettingsCancel.Location = new System.Drawing.Point(102, 114);
            this.bEmailSettingsCancel.Name = "bEmailSettingsCancel";
            this.bEmailSettingsCancel.Size = new System.Drawing.Size(75, 23);
            this.bEmailSettingsCancel.TabIndex = 11;
            this.bEmailSettingsCancel.Text = "Cancel";
            this.bEmailSettingsCancel.UseVisualStyleBackColor = true;
            this.bEmailSettingsCancel.Click += new System.EventHandler(this.bEmailSettingsCancel_Click);
            // 
            // lFromEmail
            // 
            this.lFromEmail.AutoSize = true;
            this.lFromEmail.Location = new System.Drawing.Point(43, 65);
            this.lFromEmail.Name = "lFromEmail";
            this.lFromEmail.Size = new System.Drawing.Size(58, 13);
            this.lFromEmail.TabIndex = 12;
            this.lFromEmail.Text = "From Email";
            // 
            // tbFromEmail
            // 
            this.tbFromEmail.Location = new System.Drawing.Point(107, 62);
            this.tbFromEmail.Name = "tbFromEmail";
            this.tbFromEmail.Size = new System.Drawing.Size(151, 20);
            this.tbFromEmail.TabIndex = 13;
            this.tbFromEmail.Text = "matt.howard@brownind.com";
            // 
            // FormSettingsEmailAccount
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(276, 154);
            this.Controls.Add(this.tbFromEmail);
            this.Controls.Add(this.lFromEmail);
            this.Controls.Add(this.bEmailSettingsCancel);
            this.Controls.Add(this.bEmailSettingSave);
            this.Controls.Add(this.tbPassword);
            this.Controls.Add(this.tbUsername);
            this.Controls.Add(this.tbExchangeServer);
            this.Controls.Add(this.lPassword);
            this.Controls.Add(this.lUsername);
            this.Controls.Add(this.lExchangeServer);
            this.Name = "FormSettingsEmailAccount";
            this.Text = "Email Account Settings";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lExchangeServer;
        private System.Windows.Forms.Label lUsername;
        private System.Windows.Forms.Label lPassword;
        private System.Windows.Forms.TextBox tbExchangeServer;
        private System.Windows.Forms.TextBox tbUsername;
        private System.Windows.Forms.TextBox tbPassword;
        private System.Windows.Forms.Button bEmailSettingSave;
        private System.Windows.Forms.Button bEmailSettingsCancel;
        private System.Windows.Forms.Label lFromEmail;
        private System.Windows.Forms.TextBox tbFromEmail;
    }
}