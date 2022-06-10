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
    public partial class FormSettingsShawTuftex : Form
    {
        private FormMain mainForm = null;
        MethodsCommon methods = new MethodsCommon();
        FormElements formElements = new FormElements();

        public FormSettingsShawTuftex(FormMain mainForm)
        {
            InitializeComponent();

            tbShawHotfolder.Text = Settings.Default.shawHotfolder;
            formElements.SettingsFilePath(tbShawHotfolder, fbdShawHotfolder);

            tbShawError.Text = Settings.Default.shawErrorFolder;
            formElements.SettingsFilePath(tbShawError, fbdShawError);

            tbShawArchive.Text = Settings.Default.shawArchiveFolder;
            formElements.SettingsFilePath(tbShawArchive, fbdShawArchive);

            tbShawPdfs.Text = Settings.Default.shawPdfs;
            formElements.SettingsFilePath(tbShawPdfs, fbdShawPdfs);

            tbHpOutput.Text = Settings.Default.shawHpOutput;
            formElements.SettingsFilePath(tbHpOutput, fbdHpOutput);

            tbIndigo5600.Text = Settings.Default.shawIndigo5600;
            formElements.SettingsFilePath(tbIndigo5600, fbdIndigo5600);

            tbIndigo6800.Text = Settings.Default.shawIndigo6800;
            formElements.SettingsFilePath(tbIndigo6800, fbdIndigo6800);

            tbXMFHotfolder.Text = Settings.Default.xmfHotfolders;
            formElements.SettingsFilePath(tbXMFHotfolder, fbdXMFHotfolder);

            tbJpgImages.Text = Settings.Default.tuftexJpg;
            formElements.SettingsFilePath(tbJpgImages, fbdJpgImages);

            tbXmlOutput.Text = Settings.Default.tuftexXmlHotfolder;
            formElements.SettingsFilePath(tbXmlOutput, fbdXmlOutput);

            tbMiscOutput.Text = Settings.Default.tuftexMiscHotfolder;
            formElements.SettingsFilePath(tbMiscOutput, fbdMiscOutput);            

            this.mainForm = mainForm;
        }

        public void bSettingSave_Click(object sender, EventArgs e)
        {
            Settings.Default.shawHotfolder = tbShawHotfolder.Text;
            Settings.Default.shawErrorFolder = tbShawError.Text;
            Settings.Default.shawArchiveFolder = tbShawArchive.Text;
            Settings.Default.shawPdfs = tbShawPdfs.Text;
            Settings.Default.shawHpOutput = tbHpOutput.Text;
            Settings.Default.shawIndigo5600 = tbIndigo5600.Text;
            Settings.Default.shawIndigo6800 = tbIndigo6800.Text;
            Settings.Default.xmfHotfolders = tbXMFHotfolder.Text;
            Settings.Default.tuftexJpg = tbJpgImages.Text;
            Settings.Default.tuftexXmlHotfolder = tbXmlOutput.Text;
            Settings.Default.tuftexMiscHotfolder = tbMiscOutput.Text;
            Settings.Default.Save();
            mainForm.rtMain.AppendText(DateTime.Now + " | Settings Change Saved. \r\n", Color.Black, FontStyle.Regular);
            FormSettings.ActiveForm.Close();
        }

        public void bSettingCancel_Click(object sender, EventArgs e)
        {            
            mainForm.rtMain.AppendText(DateTime.Now + " | Settings Change Cancelled. \r\n", Color.Black, FontStyle.Regular);
            FormSettings.ActiveForm.Close();
        }

        private void bShawHotfolder_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbShawHotfolder, fbdShawHotfolder);
        }

        private void bShawError_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbShawError, fbdShawError);
        }

        private void bShawArchive_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbShawArchive, fbdShawArchive);
        }

        private void bShawPdfs_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbShawPdfs, fbdShawPdfs);
        }

        private void bHpOutput_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbHpOutput, fbdHpOutput);
        }

        private void bIndigo5600_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbIndigo5600, fbdIndigo5600);
        }

        private void bIndigo6800_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbIndigo6800, fbdIndigo6800);
        }  

        private void bXMFHotfolder_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbXMFHotfolder, fbdXMFHotfolder);
        }        

        private void bJpgImages_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbJpgImages, fbdJpgImages);
        }

        private void bXmlOutput_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbXmlOutput, fbdXmlOutput);
        }

        private void bMiscOuput_Click(object sender, EventArgs e)
        {
            formElements.SettingsClick(tbMiscOutput, fbdMiscOutput);
        }        

        private void bXmlLabelSettings_Click(object sender, EventArgs e)
        {
            FileInfo fi = new FileInfo("Variables.xml");
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start("Variables.xml");
            }
            else
            {
                MessageBox.Show("File Does Not Exist.  Ask Matt");
            }
        }
        private void bMiscLabelSettings_Click(object sender, EventArgs e)
        {
            FileInfo fi = new FileInfo(Brown_Prepress_Automation.FormMain.Globals.appDir + "\\Type\\labels.xls");
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(Brown_Prepress_Automation.FormMain.Globals.appDir + "\\Type\\labels.xls");
            }
            else
            {
                MessageBox.Show("File Does Not Exist.  Ask Matt");
            }
        }

        private void bEmail_Click(object sender, EventArgs e)
        {
            FormSettingsShawEmail formSettingsShawEmail = new FormSettingsShawEmail();
            formSettingsShawEmail.Show();
        } 
    }
}
