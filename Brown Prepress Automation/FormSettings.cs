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
using Microsoft.Win32;

namespace Brown_Prepress_Automation
{
    public partial class FormSettings : Form
    {
        private FormMain mainForm = null;
        MethodsCommon methods = new MethodsCommon();
        FormElements formElements = new FormElements();

        public FormSettings(FormMain mainForm)
        {
            InitializeComponent();

            tbHotfolder.Text = Settings.Default.hotFolder;
            formElements.SettingsFilePath(tbHotfolder, fbdHotfolder);

            tbErrorFolder.Text = Settings.Default.errorFolder;
            formElements.SettingsFilePath(tbErrorFolder, fbdErrorFolder);

            tbArchiveFolder.Text = Settings.Default.archiveFolder;
            formElements.SettingsFilePath(tbArchiveFolder, fbdArchiveFolder);

            tbBlueline.Text = Settings.Default.bluelineOutput;
            formElements.SettingsFilePath(tbBlueline, fbdBlueline);

            tbParts.Text = Settings.Default.parts;
            formElements.SettingsFilePath(tbParts, fbdParts);

            tbTemp.Text = Settings.Default.tempDir;
            formElements.SettingsFilePath(tbTemp, fbdTemp);

            cbAutoUpdate.Checked = Settings.Default.updateCheck;

            cbSendEmails.Checked = Settings.Default.sendEmails;

            cbDebugOn.Checked = Settings.Default.debugOn;

            cbPrinters.Text = Settings.Default.printer;

            string listOfPrinters;
            for (int i = 0; i < PrinterSettings.InstalledPrinters.Count; i++)
            {
                listOfPrinters = PrinterSettings.InstalledPrinters[i];
                cbPrinters.Items.Add(listOfPrinters);
            }

            this.mainForm = mainForm;
        }

        public void bSettingSave_Click(object sender, EventArgs e)
        {
            Settings.Default.hotFolder = tbHotfolder.Text;         
            Settings.Default.errorFolder = tbErrorFolder.Text;
            Settings.Default.archiveFolder = tbArchiveFolder.Text; 
            Settings.Default.bluelineOutput = tbBlueline.Text;
            Settings.Default.parts = tbParts.Text;
            Settings.Default.tempDir = tbTemp.Text; 
            Settings.Default.printer = cbPrinters.Text; 
            Settings.Default.Save();
            mainForm.rtMain.AppendText(DateTime.Now + " | Settings Change Saved. \r\n", Color.Black, FontStyle.Regular);
            FormSettings.ActiveForm.Close();
        }

        public void bSettingCancel_Click(object sender, EventArgs e)
        {            
            mainForm.rtMain.AppendText(DateTime.Now + " | Settings Change Cancelled. \r\n", Color.Black, FontStyle.Regular);
            FormSettings.ActiveForm.Close();
        }

        private void bHotfolder_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbHotfolder, fbdHotfolder);
        }        

        private void bErrorFolder_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbErrorFolder, fbdErrorFolder);
        }

        private void bArchiveFolder_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbArchiveFolder, fbdArchiveFolder);
        }

        private void bBlueline_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbBlueline, fbdBlueline);
        }

        private void bParts_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbParts, fbdParts);
        }

        private void bTemp_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbTemp, fbdTemp);
        }

        private void cbAutoUpdate_CheckedChanged(object sender, EventArgs e)
        {
            if (cbAutoUpdate.Checked == true)
            {
                Settings.Default.updateCheck = true;
            }
            else
            {
                Settings.Default.updateCheck = false;
            }
        }

        private void cbSendEmails_CheckedChanged(object sender, EventArgs e)
        {
            if (cbSendEmails.Checked == true)
            {
                Settings.Default.sendEmails = true;
            }
            else
            {
                Settings.Default.sendEmails = false;
            }
        }

        private void cbDebugOn_CheckedChanged(object sender, EventArgs e)
        {
            if (cbDebugOn.Checked == true)
            {
                Settings.Default.debugOn = true;
            }
            else
            {
                Settings.Default.debugOn = false;
            }
        }

        private void bEmailSettings_Click(object sender, EventArgs e)
        {
            FormSettingsEmailAccount formSettingsEmail = new FormSettingsEmailAccount();
            formSettingsEmail.Show();
        }

        private void cbPrinters_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (cbPrinters.SelectedIndex != -1)
            {
                Settings.Default.printer = cbPrinters.Text;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FormSettingsShawTuftex formSettingsTuftex = new FormSettingsShawTuftex(mainForm);
            formSettingsTuftex.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Settings.Default.Reset();
            //methods.SettingsDelete();
            FormSettings.ActiveForm.Close();
        }

        private void bArmstrong_Click(object sender, EventArgs e)
        {
            FormSettingsArmstrong formSettingsArmstrong = new FormSettingsArmstrong(mainForm);
            formSettingsArmstrong.Show();
        }

        private void bNourison_Click(object sender, EventArgs e)
        {
            FormSettingsNourison formSettingsNourison = new FormSettingsNourison(mainForm);
            formSettingsNourison.Show();
        }               
    }
}
