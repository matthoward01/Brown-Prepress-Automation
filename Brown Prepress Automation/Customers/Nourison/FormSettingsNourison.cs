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
    public partial class FormSettingsNourison : Form
    {
        private FormMain mainForm = null;
        MethodsCommon methods = new MethodsCommon();
        FormElements formElements = new FormElements();

        public FormSettingsNourison(FormMain mainForm)
        {
            InitializeComponent();

            tbpdfs.Text = Settings.Default.nourisonPdfs;
            formElements.SettingsFilePath(tbpdfs, fbdPdfs);

            tbHotfolder.Text = Settings.Default.nourisonHotfolder;
            formElements.SettingsFilePath(tbHotfolder, fbdHotfolder);

            tbError.Text = Settings.Default.nourisonErrorFolder;
            formElements.SettingsFilePath(tbError, fbdError);

            tbArchive.Text = Settings.Default.nourisonArchiveFolder;
            formElements.SettingsFilePath(tbArchive, fbdArchive);

            tbHpOutput.Text = Settings.Default.nourisonHpOutput;
            formElements.SettingsFilePath(tbHpOutput, fbdHpOutput);

            if (Settings.Default.nourisonCommon == true)
            {
                cbCommon.Checked = true;
            }
            else
            {
                cbCommon.Checked = false;
            }
  
            this.mainForm = mainForm;
        }

        public void bSettingSave_Click(object sender, EventArgs e)
        {
            Settings.Default.nourisonPdfs = tbpdfs.Text;
            Settings.Default.nourisonHotfolder = tbHotfolder.Text;
            Settings.Default.nourisonErrorFolder = tbError.Text;
            Settings.Default.nourisonArchiveFolder = tbArchive.Text;
            Settings.Default.nourisonHpOutput = tbHpOutput.Text;
            if (cbCommon.Checked == true)
            {
                Settings.Default.nourisonCommon = true;
            }
            else
            {
                Settings.Default.nourisonCommon = false;
            }
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

        private void bHpOutput_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbHpOutput, fbdHpOutput);
        }
    }
}
