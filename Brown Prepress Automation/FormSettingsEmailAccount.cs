using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Brown_Prepress_Automation.Properties;

namespace Brown_Prepress_Automation
{
    public partial class FormSettingsEmailAccount : Form
    {
        public FormSettingsEmailAccount()
        {
            InitializeComponent();

            tbExchangeServer.Text = Settings.Default.exchangeServer;
            tbFromEmail.Text = Settings.Default.fromEmail;
            tbUsername.Text = Settings.Default.username;
            tbPassword.Text = Settings.Default.password;
        }

        private void bEmailSettingsCancel_Click(object sender, EventArgs e)
        {
            FormSettings.ActiveForm.Close();
        }

        private void bEmailSettingSave_Click(object sender, EventArgs e)
        {
            Settings.Default.exchangeServer = tbExchangeServer.Text;
            Settings.Default.fromEmail = tbFromEmail.Text;
            Settings.Default.username = tbUsername.Text;
            Settings.Default.password = tbPassword.Text;
            Settings.Default.Save();
            FormSettings.ActiveForm.Close();
        }
    }
}
