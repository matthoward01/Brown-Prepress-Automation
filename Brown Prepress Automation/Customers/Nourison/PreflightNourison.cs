using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelLibrary.SpreadSheet;
using Brown_Prepress_Automation.Properties;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace Brown_Prepress_Automation
{
    class PreflightNourison
    {
        NourisonOutput nourisonOutput = new NourisonOutput();
        MethodsCommon methods = new MethodsCommon();
        MethodsMySQL methodsMySql = new MethodsMySQL();
        bool errorCheck = false;        

        public void PreflightNourisonPop(FormMain mainForm, string passedFile)
        {
            List<string> exceptionErrorList = new List<string>();

            try
            {
                List<string> popFileNameList = new List<string>();
                //List<string> popList = new List<string>();
                string workingFile = passedFile;
                string popType = "nourison";
                string popSubFolder = "Nourison";
                exceptionErrorList.Add(workingFile);
                Workbook book = Workbook.Load(Settings.Default.nourisonHotfolder + "\\" + workingFile);
                Worksheet sheet = book.Worksheets[0];
                string ordernumber = sheet.Cells[0, 11].StringValue;
                if (sheet.Cells[0, 0].StringValue.Contains("IDG"))
                {
                    popType = "idg";
                    popSubFolder = "IDG";
                }
                else if (sheet.Cells[0, 0].StringValue.Contains("CK"))
                {
                    popType = "ck";
                }

                popFileNameList.AddRange(NourisonCheck(mainForm, sheet, workingFile, popSubFolder, 9, 74, 1, 2, 4));

                if (popType != "ck")
                {
                    popFileNameList.AddRange(NourisonCheck(mainForm, sheet, workingFile, popSubFolder, 9, 74, 6, 7, 9));
                    popFileNameList.AddRange(NourisonCheck(mainForm, sheet, workingFile, popSubFolder, 9, 74, 11, 12, 14));
                    popFileNameList.AddRange(NourisonCheck(mainForm, sheet, workingFile, popSubFolder, 78, 156, 1, 2, 4));
                    popFileNameList.AddRange(NourisonCheck(mainForm, sheet, workingFile, popSubFolder, 78, 156, 6, 7, 9));
                    popFileNameList.AddRange(NourisonCheck(mainForm, sheet, workingFile, popSubFolder, 78, 156, 11, 12, 14));
                    popFileNameList.AddRange(NourisonCheck(mainForm, sheet, workingFile, popSubFolder, 160, 238, 1, 2, 4));
                    popFileNameList.AddRange(NourisonCheck(mainForm, sheet, workingFile, popSubFolder, 160, 238, 6, 7, 9));
                    popFileNameList.AddRange(NourisonCheck(mainForm, sheet, workingFile, popSubFolder, 160, 238, 11, 12, 14));
                }
                if (sheet.Cells[0, 10].StringValue.Contains("4pg"))
                {
                    popFileNameList.AddRange(NourisonCheck(mainForm, sheet, workingFile, popSubFolder, 242, 318, 1, 2, 4));
                    popFileNameList.AddRange(NourisonCheck(mainForm, sheet, workingFile, popSubFolder, 242, 318, 6, 7, 9));
                    popFileNameList.AddRange(NourisonCheck(mainForm, sheet, workingFile, popSubFolder, 242, 318, 11, 12, 14));
                }
                Random rnd = new Random();
                int common = rnd.Next(0, 4);
                while (popFileNameList.Count % 10 != 0)
                {
                    if (Settings.Default.nourisonCommon == true)
                    {
                        string[] filler = { "Intro DiningRoom.pdf", "Intro LivingRoom.pdf", "Intro Bedroom.pdf", "Intro CustomCrafted.pdf" };
                        popFileNameList.Add(Settings.Default.nourisonPdfs + "\\" + popSubFolder + "\\" + filler[common]);
                        common = rnd.Next(0, 4);
                    }
                    else
                    {
                        string[] filler = { "Blank.pdf", "Blank.pdf", "Blank.pdf", "Blank.pdf" };
                        popFileNameList.Add(Settings.Default.nourisonPdfs + "\\" + popSubFolder + "\\" + filler[common]);
                        common = rnd.Next(0, 4);
                    }
                }
                if (errorCheck)
                {
                    DialogResult dialogResult = MessageBox.Show("Continue with errors?", "Nourison POP Warning", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        throw new Exception("Processing of Nourison Sheet has been cancelled.");
                    }
                }
                if (workingFile.ToLower().Contains("check"))
                {

                }
                else
                {
                    nourisonOutput.pdfNourisonPOP(mainForm, popFileNameList, ordernumber);
                    mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + " | " + Path.GetFileNameWithoutExtension(workingFile) + " is done.\r\n", Color.Black, FontStyle.Regular); }));
                    popFileNameList.Clear();
                }
                exceptionErrorList.Remove(workingFile);

                methodsMySql.InsertOrders("nourison", Path.GetFileNameWithoutExtension(workingFile), "", "", "", "", "", "", "");

                //Archive

                if (File.Exists(Settings.Default.nourisonArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + passedFile))
                {
                    System.IO.File.Delete(Settings.Default.nourisonArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + passedFile);
                }
                System.IO.Directory.CreateDirectory(Settings.Default.nourisonArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\");
                if (File.Exists(Settings.Default.nourisonHotfolder + "\\" + passedFile))
                {
                    System.IO.File.Copy(Settings.Default.nourisonHotfolder + "\\" + passedFile, Settings.Default.nourisonArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + passedFile);
                }
                if (File.Exists(Settings.Default.nourisonHotfolder + "\\" + passedFile))
                {
                    System.IO.File.Delete(Settings.Default.nourisonHotfolder + "\\" + passedFile);
                }
            }
            catch (Exception ex)
            {
                foreach (string e in exceptionErrorList)
                {
                    using (System.IO.StreamWriter errorFile = new System.IO.StreamWriter(Settings.Default.nourisonErrorFolder + "\\" + e + ".txt", true))
                    {
                        errorFile.WriteLine(DateTime.Now + " | " + ex.Message + "\r\nCheck your spreadsheet format. \r\nAsk Matt with questions.");
                        Console.WriteLine(DateTime.Now + " | " + ex.Message + "\r\nCheck your spreadsheet format. \r\nAsk Matt with questions.");
                        mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + ex.Message + "\r\nCheck your spreadsheet format. \r\nAsk Matt with questions.\r\n", Color.Red, FontStyle.Regular); }));
                    }
                    if (File.Exists(Settings.Default.nourisonErrorFolder + "\\" + e) && File.Exists(Settings.Default.nourisonHotfolder + "\\" + e))
                    {
                        System.IO.File.Delete(Settings.Default.nourisonErrorFolder + "\\" + e);
                    }
                    if (File.Exists(Settings.Default.nourisonHotfolder + "\\" + e))
                    {
                        System.IO.File.Move(Settings.Default.nourisonHotfolder + "\\" + e, Settings.Default.nourisonErrorFolder + "\\" + e);
                    }
                }
            }
            finally
            {
                exceptionErrorList.Clear();
            }
        }

        public List<string> NourisonCheck(FormMain mainForm, Worksheet sheet, string workingFile, string popSubFolder, int startY, int endY, int designX, int colorX, int qtyX)
        {
            List<string> popFileNameList = new List<string>();
            List<int> popQtyList = new List<int>();
            

            for (int i = startY; i <= endY; i++)
            {
                string popName = "";
                int popQty = 0;
                if (sheet.Cells[i, qtyX].StringValue != "")
                {
                    popQty = Convert.ToInt32(sheet.Cells[i, qtyX].StringValue);
                    popName = sheet.Cells[i, designX].StringValue.Trim() + " " + sheet.Cells[i, colorX].StringValue.Trim();


                    popName = popName.Replace("_LR", "");
                    if (!popName.Contains(".pdf"))
                    {
                        popName = popName + ".pdf";
                    }

                    string[] popFiles = Directory.GetFiles(Settings.Default.nourisonPdfs + "\\" + popSubFolder + "\\", popName, SearchOption.AllDirectories);

                    if (popFiles == null || popFiles.Length < 1)
                    {
                        errorCheck = true;
                        using (System.IO.StreamWriter errorFile = new System.IO.StreamWriter(Settings.Default.nourisonErrorFolder + "\\" + workingFile + ".txt", true))
                            errorFile.WriteLine(DateTime.Now + " | " + popName + " is missing. Please Fix and resubmit spreadsheet via the hotfolder.");
                        mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + popName + " is missing from " + workingFile + "... \r\n", Color.Red, FontStyle.Regular); }));
                    }

                    while (popQty != 0)
                    {
                        foreach (string s in popFiles)
                        {
                            popFileNameList.Add(s.Trim());
                        }
                        popQty--;
                    }
                }
            }           

            return popFileNameList;
        }
    }
}
