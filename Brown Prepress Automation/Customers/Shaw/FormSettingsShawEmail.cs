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
    public partial class FormSettingsShawEmail : Form
    {
        public FormSettingsShawEmail()
        {
            InitializeComponent();

            foreach (string email in Settings.Default.shawEmailList)
            {
                tbShawEmailList.AppendText(email.Trim()+"\r\n");
            }
            foreach (string email in Settings.Default.tuftexEmailList)
            {
                tbTuftexEmailList.AppendText(email.Trim() + "\r\n");
            }
            tbShawEmailList.Text.Trim();
            tbTuftexEmailList.Text.Trim();
        }

        private void bEmailSettingsCancel_Click(object sender, EventArgs e)
        {
            FormSettings.ActiveForm.Close();
        }

        private void bEmailSettingSave_Click(object sender, EventArgs e)
        {   Settings.Default.shawEmailList.Clear();
            Settings.Default.tuftexEmailList.Clear();
            string[] modifiedShawEmailList = tbShawEmailList.Text.Split('\n');
            string[] modifiedTuftexEmailList = tbTuftexEmailList.Text.Split('\n');
            foreach (string email in modifiedShawEmailList)
            {
                if (email != "")
                {
                    Settings.Default.shawEmailList.Add(email.Trim());
                }
            }
            foreach (string email in modifiedTuftexEmailList)
            {
                if (email != "")
                {
                    Settings.Default.tuftexEmailList.Add(email.Trim());
                }
            }
            Settings.Default.Save();
            FormSettings.ActiveForm.Close();
        }
    }
}
