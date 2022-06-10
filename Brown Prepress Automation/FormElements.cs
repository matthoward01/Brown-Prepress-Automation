using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Brown_Prepress_Automation.Properties;
using System.Windows.Forms;
using System.IO;

namespace Brown_Prepress_Automation
{
    class FormElements
    {
        public void SettingsFilePath(TextBox textBox, FolderBrowserDialog folderBrowserDialog)
        {
            if (textBox.Text == "")
            {
                textBox.Text = folderBrowserDialog.SelectedPath;
            }
        }

        public void SettingsClick(TextBox textBox, FolderBrowserDialog folderBrowserDialog)
        {
            if (Directory.Exists(textBox.Text))
            {
                folderBrowserDialog.SelectedPath = textBox.Text;
            }
            else
            {
                folderBrowserDialog.SelectedPath = Settings.Default.lastFolder;
            }

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                textBox.Text = folderBrowserDialog.SelectedPath;
                Settings.Default.lastFolder = folderBrowserDialog.SelectedPath;
            }
        }

    }
}
