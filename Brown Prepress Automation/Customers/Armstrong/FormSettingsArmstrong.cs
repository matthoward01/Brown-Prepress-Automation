using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Brown_Prepress_Automation.Properties;
using System.Drawing.Printing;

namespace Brown_Prepress_Automation
{
    public partial class FormSettingsArmstrong : Form
    {
        private FormMain mainForm = null;
        MethodsCommon methods = new MethodsCommon();
        FormElements formElements = new FormElements();

        public FormSettingsArmstrong(FormMain mainForm)
        {
            InitializeComponent();

            tbAHFpdfs.Text = Settings.Default.armstrongAHFpdfs;
            formElements.SettingsFilePath(tbAHFpdfs, fbdAHFpdfs);

            tbpdfs.Text = Settings.Default.armstrongPdfs;
            formElements.SettingsFilePath(tbpdfs, fbdPdfs);

            tbHotfolder.Text = Settings.Default.armstrongHotfolder;
            formElements.SettingsFilePath(tbHotfolder, fbdHotfolder);

            tbError.Text = Settings.Default.armstrongErrorFolder;
            formElements.SettingsFilePath(tbError, fbdError);

            tbArchive.Text = Settings.Default.armstrongArchiveFolder;
            formElements.SettingsFilePath(tbArchive, fbdArchive);

            tbXMF5x1_75.Text = Settings.Default.XMF5x1_75;
            formElements.SettingsFilePath(tbXMF5x1_75, fbdXMF5x1_75);

            tbXMF16x16.Text = Settings.Default.XMF16x16;
            formElements.SettingsFilePath(tbXMF16x16, fbdXMF16x16);
  
            this.mainForm = mainForm;
        }

        public void bSettingSave_Click(object sender, EventArgs e)
        {
            Settings.Default.armstrongPdfs = tbpdfs.Text;
            Settings.Default.armstrongAHFpdfs = tbAHFpdfs.Text;
            Settings.Default.armstrongHotfolder = tbHotfolder.Text;
            Settings.Default.armstrongErrorFolder = tbError.Text;
            Settings.Default.armstrongArchiveFolder = tbArchive.Text;
            Settings.Default.XMF5x1_75 = tbXMF5x1_75.Text;
            Settings.Default.XMF16x16 = tbXMF16x16.Text;
            Settings.Default.Save();
            mainForm.rtMain.AppendText(DateTime.Now + " | Settings Change Saved. \r\n", Color.Black, FontStyle.Regular);
            FormSettings.ActiveForm.Close();
        }

        public void bSettingCancel_Click(object sender, EventArgs e)
        {            
            mainForm.rtMain.AppendText(DateTime.Now + " | Settings Change Cancelled. \r\n", Color.Black, FontStyle.Regular);
            FormSettings.ActiveForm.Close();
        }

        private void bPdfs_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbpdfs, fbdPdfs);
        }

        private void bHotfolder_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbHotfolder, fbdHotfolder);
        }

        private void bError_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbError, fbdError);
        }

        private void bArchive_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbArchive, fbdArchive);
        }

        private void bXMF5x1_75_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbXMF5x1_75, fbdXMF5x1_75);
        }

        private void bAHFpdfs_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbAHFpdfs, fbdAHFpdfs);
        }

        private void bXMF16x16_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbXMF16x16, fbdXMF16x16);
        }

        private void bArmstrongEmailSettings_Click(object sender, EventArgs e)
        {
            pArmstrongEmail.Visible = true;
            tbArmstrongEmail.Clear();
            foreach (string email in Settings.Default.armstrongEmailList)
            {
                tbArmstrongEmail.AppendText(email.Trim() + "\r\n");
            }
            tbArmstrongEmail.Text.Trim();
        }

        private void bArmstrongEmailCancel_Click(object sender, EventArgs e)
        {
            pArmstrongEmail.Visible = false;
        }

        private void bArmstrongEmailSave_Click(object sender, EventArgs e)
        {
            Settings.Default.armstrongEmailList.Clear();
            string[] modifiedArmstrongEmailList = tbArmstrongEmail.Text.Split('\n');
            foreach (string email in modifiedArmstrongEmailList)
            {
                if (email != "")
                {
                    Settings.Default.armstrongEmailList.Add(email.Trim());
                }
            }
            Settings.Default.Save();
            pArmstrongEmail.Visible = false;
        }
    }
}
