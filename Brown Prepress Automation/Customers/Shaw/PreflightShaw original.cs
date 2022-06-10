using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelLibrary.SpreadSheet;
using Brown_Prepress_Automation.Properties;
using System.IO;
using System.Drawing;

namespace Brown_Prepress_Automation
{
    class PreflightShaw
    {
        MethodsCommon methods = new MethodsCommon();
        OutputHPShaw outputHPShaw = new OutputHPShaw();
        OutputIndigo6800Shaw outputIndigo6800Shaw = new OutputIndigo6800Shaw();
        MethodsTicket ticket = new MethodsTicket();
        MethodsMail mail = new MethodsMail();
        PreflightPdf preflightPdf = new PreflightPdf();
        PdfProcessing pdfProcessing = new PdfProcessing();
        MethodsMySQL methodsMySQL = new MethodsMySQL();

        public void Preflight(FormMain mainForm, string passedFile)
        {
            if (passedFile.ToLower().Contains("board"))
            {
                PreflightShawBoardXLS(mainForm, passedFile);
            }
            else if (passedFile.ToLower().Contains("xml"))
            {
                //preflightTuftexXml.TuftexRun(mainForm, passedFile);
            }
            else if (passedFile.ToLower().Contains("misc") || passedFile.ToLower().Contains("mill") || passedFile.ToLower().Contains("ddp"))
            {
                //preflightTuftexMiscOld.TuftexMisOldRun(mainForm, passedFile);
            }
            else if (passedFile.ToLower().Contains("get"))
            {
                //PreflightXmlDownload(mainForm, passedFile);
            }
            else
            {
                PreflightShawLabelsXLS(mainForm, passedFile);
            }
        }

        public void PreflightShawBoardXLS(FormMain mainForm, string passedFile)
        {
            List<string> exceptionErrorList = new List<string>();

            try
            {
                List<string> coverList = new List<string>();
                List<string> linerList = new List<string>();
                List<string> flList = new List<string>();
                List<string> blList = new List<string>();
                List<string> partNumberList = new List<string>();
                List<string> coverLinerList = new List<string>();
                List<string> boardList = new List<string>();

                string workingFile = passedFile;
                exceptionErrorList.Add(workingFile);
                bool errorCheck = false;
                Workbook book = Workbook.Load(Settings.Default.shawHotfolder + "\\" + workingFile);
                Worksheet sheet = book.Worksheets[0];
                int validCellsCheck = methods.countValidCells(Settings.Default.shawHotfolder + "\\" + workingFile, 1, 0);                
                for (int i = 1; i < validCellsCheck; i++)
                {
                    string partNumber = sheet.Cells[i, 0].StringValue.Trim();
                    string coverName = sheet.Cells[i, 1].StringValue.Trim();
                    string linerName = sheet.Cells[i, 2].StringValue.Trim();
                    string flName = sheet.Cells[i, 3].StringValue.Trim();
                    string blName = sheet.Cells[i, 4].StringValue.Trim();

                    if (flName == "" || flName == " ")
                    {
                        flName = "none";
                    }
                    if (blName == "" || blName == " ")
                    {
                        blName = "none";
                    }

                    coverName = coverName.Replace("_LR", "");
                    if (!coverName.Contains(".pdf"))
                    {
                        coverName = coverName + ".pdf";
                    }
                    linerName = linerName.Replace("_LR", "");
                    if (!linerName.Contains(".pdf"))
                    {
                        linerName = linerName + ".pdf";
                    }
                    flName = flName.Replace("_LR", "");
                    if (!flName.Contains(".pdf"))
                    {
                        flName = flName + ".pdf";
                    }
                    blName = blName.Replace("_LR", "");
                    if (!blName.Contains(".pdf"))
                    {
                        blName = blName + ".pdf";
                    }
                    string[] coverFiles = Directory.GetFiles(Settings.Default.shawPdfs, coverName, SearchOption.AllDirectories);
                    string[] linerFiles = Directory.GetFiles(Settings.Default.shawPdfs, linerName, SearchOption.AllDirectories);
                    string[] flFiles = Directory.GetFiles(Settings.Default.shawPdfs, flName, SearchOption.AllDirectories);
                    string[] blFiles = Directory.GetFiles(Settings.Default.shawPdfs, blName, SearchOption.AllDirectories);

                    if (coverFiles == null || coverFiles.Length < 1)
                    {
                        errorCheck = true;
                        using (System.IO.StreamWriter errorFile = new System.IO.StreamWriter(Settings.Default.shawErrorFolder + "\\" + workingFile + ".txt", true))
                            errorFile.WriteLine(DateTime.Now + "| " + coverName + " is missing. Please Fix and resubmit spreadsheet via the hotfolder.");
                        mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + coverName + " is missing from " + passedFile + "... \r\n", Color.Red, FontStyle.Regular); }));
                    }
                    if (linerFiles == null || linerFiles.Length < 1)
                    {
                        errorCheck = true;
                        using (System.IO.StreamWriter errorFile = new System.IO.StreamWriter(Settings.Default.shawErrorFolder + "\\" + workingFile + ".txt", true))
                            errorFile.WriteLine(DateTime.Now + "| " + linerName + " is missing. Please Fix and resubmit spreadsheet via the hotfolder.");
                        mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + linerName + " is missing from " + passedFile + "... \r\n", Color.Red, FontStyle.Regular); }));
                    }
                    if (flFiles == null || flFiles.Length < 1)
                    {
                        errorCheck = true;
                        using (System.IO.StreamWriter errorFile = new System.IO.StreamWriter(Settings.Default.shawErrorFolder + "\\" + workingFile + ".txt", true))
                            errorFile.WriteLine(DateTime.Now + "| " + flName + " is missing. Please Fix and resubmit spreadsheet via the hotfolder.");
                        mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + flName + " is missing from " + passedFile + "... \r\n", Color.Red, FontStyle.Regular); }));
                    }
                    if (blFiles == null || blFiles.Length < 1)
                    {
                        errorCheck = true;
                        using (System.IO.StreamWriter errorFile = new System.IO.StreamWriter(Settings.Default.shawErrorFolder + "\\" + workingFile + ".txt", true))
                            errorFile.WriteLine(DateTime.Now + "| " + blName + " is missing. Please Fix and resubmit spreadsheet via the hotfolder.");
                        mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + blName + " is missing from " + passedFile + "... \r\n", Color.Red, FontStyle.Regular); }));
                    }

                    foreach (string s in coverFiles)
                    {
                        coverList.Add(s.Trim());
                    }
                    foreach (string s in coverFiles)
                    {
                        boardList.Add(s.Trim());
                    }
                    foreach (string s in linerFiles)
                    {
                        linerList.Add(s.Trim());
                    }
                    foreach (string s in flFiles)
                    {
                        flList.Add(s.Trim());
                    }
                    foreach (string s in blFiles)
                    {
                        blList.Add(s.Trim());
                    }

                    partNumberList.Add(partNumber.Trim());
                }
                if (errorCheck == true)
                {
                    if (File.Exists(Settings.Default.shawErrorFolder + "\\" + workingFile) && File.Exists(Settings.Default.shawHotfolder + "\\" + workingFile))
                    {
                        System.IO.File.Delete(Settings.Default.shawErrorFolder + "\\" + workingFile);
                    }
                    if (File.Exists(Settings.Default.shawHotfolder + "\\" + workingFile))
                    {
                        System.IO.File.Move(Settings.Default.shawHotfolder + "\\" + workingFile, Settings.Default.shawErrorFolder + "\\" + workingFile);
                    }
                    if (Settings.Default.sendEmails == true)
                    {
                        mail.SendMailShawTeam(workingFile, true);
                    }
                }

                else
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    int fileProgressStep = (int)Math.Ceiling(((double)100) / boardList.Count);

                    foreach (string board in boardList)
                    {
                        string productType = "";
                        if (partNumberList[0].Contains("CAP"))
                        {
                            productType = "CAP";
                        }
                        else if (partNumberList[0].Contains("CARD"))
                        {
                            productType = "CARD";
                        }
                        else if (partNumberList[0].Contains("PCS"))
                        {
                            productType = "PCS";
                        }
                        else if (partNumberList[0].Contains("PROTO"))
                        {
                            productType = "PROTO";
                        }
                        else
                        {
                            throw new Exception("Product Type is not supported.  Should be CAP, CARD, PCS, or PROTO");
                        }
                        if (board.Contains("717") || board.Contains("758") || board.Contains("J75") || board.Contains("724") || board.Contains("J73") || board.Contains("J77") || board.Contains("Supplied Boards") || board.Contains("J39") || board.Contains("320") || board.Contains("458"))
                        {
                            System.IO.Directory.CreateDirectory(Settings.Default.parts + "\\" + productType + "\\" + partNumberList[0].Substring(0, 8).Replace(productType + "-", "") + "000" + "\\" + partNumberList[0] + "\\");
                            System.IO.File.Copy(coverList[0], Settings.Default.parts + "\\" + productType + "\\" + partNumberList[0].Substring(0, 8).Replace(productType + "-", "") + "000" + "\\" + partNumberList[0] + "\\" + partNumberList[0] + " Cover.pdf", true);
                            System.IO.File.Copy(linerList[0], Settings.Default.parts + "\\" + productType + "\\" + partNumberList[0].Substring(0, 8).Replace(productType + "-", "") + "000" + "\\" + partNumberList[0] + "\\" + partNumberList[0] + " Liner.pdf", true);
                            string[] coverLinerArray = outputHPShaw.pdf1upBoard(partNumberList[0], coverList[0], linerList[0], flList[0], blList[0]);
                            System.IO.File.Copy(Settings.Default.tempDir + "\\" + partNumberList[0] + " HP 40x55 - Printable.pdf", Settings.Default.parts + "\\" + productType + "\\" + partNumberList[0].Substring(0, 8).Replace(productType + "-", "") + "000" + "\\" + partNumberList[0] + "\\" + partNumberList[0] + " HP 40x55 - Printable.pdf", true);
                            coverLinerList = coverLinerArray.ToList();
                            if (Settings.Default.debugOn != true)
                            {
                                if (!board.Contains("Supplied Boards"))
                                {
                                    System.IO.File.Copy(Settings.Default.tempDir + "\\" + partNumberList[0] + " HP 40x55 - Printable.pdf", Settings.Default.bluelineOutput + "\\" + partNumberList[0] + " HP 40x55 - Printable.pdf", true);
                                }
                                else
                                {
                                    foreach (string f in coverLinerList)
                                    {
                                        System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(f) + ".pdf", Settings.Default.shawHotfolder + "\\" + Path.GetFileNameWithoutExtension(f) + ".pdf", true);
                                    }
                                }
                                methodsMySQL.InsertPrepressLogAutomation(partNumberList[0]);
                                System.IO.File.Copy(Settings.Default.tempDir + "\\" + partNumberList[0] + " HP 40x55 - Printable.pdf", Settings.Default.shawHpOutput + "\\" + partNumberList[0] + " HP 40x55 - Printable.pdf", true);
                            }

                            foreach (string f in coverLinerList)
                            {
                                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(f) + ".pdf", Settings.Default.parts + "\\" + productType + "\\" + partNumberList[0].Substring(0, 8).Replace(productType + "-", "") + "000" + "\\" + partNumberList[0] + "\\" + Path.GetFileNameWithoutExtension(f) + ".pdf", true);
                            }

                            
                            partNumberList.RemoveAt(0);
                            coverList.RemoveAt(0);
                            linerList.RemoveAt(0);
                            flList.RemoveAt(0);
                            blList.RemoveAt(0);
                            mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.Step = fileProgressStep; }));
                            mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                        }                        
                        else
                        {
                            System.IO.Directory.CreateDirectory(Settings.Default.parts + "\\" + productType + "\\" + partNumberList[0].Substring(0, 8).Replace(productType + "-", "") + "000" + "\\" + partNumberList[0] + "\\");
                            System.IO.File.Copy(coverList[0], Settings.Default.parts + "\\" + productType + "\\" + partNumberList[0].Substring(0, 8).Replace(productType + "-", "") + "000" + "\\" + partNumberList[0] + "\\" + partNumberList[0] + " Cover.pdf", true);
                            System.IO.File.Copy(linerList[0], Settings.Default.parts + "\\" + productType + "\\" + partNumberList[0].Substring(0, 8).Replace(productType + "-", "") + "000" + "\\" + partNumberList[0] + "\\" + partNumberList[0] + " Liner.pdf", true);
                            if (Settings.Default.debugOn == false)
                            {
                                System.IO.File.Copy(coverList[0], Settings.Default.bluelineOutput + "\\" + partNumberList[0] + " Cover.pdf", true);
                                System.IO.File.Copy(linerList[0], Settings.Default.bluelineOutput + "\\" + partNumberList[0] + " Liner.pdf", true);
                            }
                            partNumberList.RemoveAt(0);
                            coverList.RemoveAt(0);
                            linerList.RemoveAt(0);
                            flList.RemoveAt(0);
                            blList.RemoveAt(0);
                            mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.Step = fileProgressStep; }));
                            mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                        }
                    }

                    coverList.Clear();
                    linerList.Clear();
                    flList.Clear();
                    blList.Clear();
                    partNumberList.Clear();

                    if (Settings.Default.debugOn == false)
                    {
                        ticket.shawBoardPrintableTicket(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile));
                    }
                    if (Settings.Default.sendEmails == true)
                    {
                        mail.SendMailShawTeam(workingFile, false);
                    }
                }
                if (!workingFile.ToLower().Contains("pdf"))
                {
                    if (File.Exists(Settings.Default.shawArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + workingFile))
                    {
                        System.IO.File.Delete(Settings.Default.shawArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + workingFile);
                    }
                    System.IO.Directory.CreateDirectory(Settings.Default.shawArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\");
                    if (File.Exists(Settings.Default.shawHotfolder + workingFile))
                    {
                        System.IO.File.Copy(Settings.Default.shawHotfolder + workingFile, Settings.Default.shawArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + workingFile);
                    }
                }
                if (File.Exists(Settings.Default.shawHotfolder + "\\" + workingFile))
                {
                    System.IO.File.Delete(Settings.Default.shawHotfolder + "\\" + workingFile);
                }
                exceptionErrorList.Remove(workingFile);
            }
            catch (Exception ex)
            {
                foreach (string e in exceptionErrorList)
                {
                    using (System.IO.StreamWriter errorFile = new System.IO.StreamWriter(Settings.Default.shawErrorFolder + "\\" + e + ".txt", true))
                    {
                        errorFile.WriteLine(DateTime.Now + "| " + ex.Message + "\r\nCheck your spreadsheet format. \r\nAsk Matt with questions.");
                        Console.WriteLine(DateTime.Now + "| " + ex.Message + "\r\nCheck your spreadsheet format. \r\nAsk Matt with questions.");
                        mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + ex.Message + "\r\nCheck your spreadsheet format. \r\nAsk Matt with questions.\r\n", Color.Red, FontStyle.Regular); }));
                        if (Settings.Default.sendEmails == true)
                        {
                            mail.SendMailShawTeam(e, true);
                        }
                    }
                    if (File.Exists(Settings.Default.shawErrorFolder + "\\" + e) && File.Exists(Settings.Default.shawHotfolder + "\\" + e))
                    {
                        System.IO.File.Delete(Settings.Default.shawErrorFolder + "\\" + e);
                    }
                    if (File.Exists(Settings.Default.shawHotfolder + "\\" + e))
                    {
                        System.IO.File.Move(Settings.Default.shawHotfolder + "\\" + e, Settings.Default.shawErrorFolder + "\\" + e);
                    }
                }
            }
            finally
            {
                exceptionErrorList.Clear();
            }
        }

        public void PreflightShawLabelsXLS(FormMain mainForm, string passedFile)
        {
            List<string> exceptionErrorList = new List<string>();
            List<string> errorEmailList = new List<string>();

            try
            {
                List<string> dropboxFileName = new List<string>();
                List<int> dropboxQty = new List<int>();
                List<string> diffPerSheet = new List<string>();
                string workingFile = passedFile;
                exceptionErrorList.Add(workingFile);
                bool errorCheck = false;
                Workbook book = Workbook.Load(Settings.Default.shawHotfolder + "\\" + workingFile);
                Worksheet sheet = book.Worksheets[0];
                int validCellsCheck = methods.countValidCells(Settings.Default.shawHotfolder + "\\" + workingFile, 1, 0);
                string previousFile = "";
                for (int i = 1; i < validCellsCheck; i++)
                {
                    string labelName = sheet.Cells[i, 0].StringValue.Trim();
                    int labelQty = Convert.ToInt32(sheet.Cells[i, 5].StringValue);

                    labelName = labelName.Replace("_LR", "");
                    if (!labelName.Contains(".pdf"))
                    {
                        labelName = labelName + ".pdf";
                    }

                    string[] labelFiles = Directory.GetFiles(Settings.Default.shawPdfs, labelName, SearchOption.AllDirectories);

                    if (labelFiles == null || labelFiles.Length < 1)
                    {
                        errorCheck = true;
                        using (System.IO.StreamWriter errorFile = new System.IO.StreamWriter(Settings.Default.shawErrorFolder + "\\" + workingFile + ".txt", true))
                            errorFile.WriteLine(DateTime.Now + "| " + labelName + " is missing.");
                        errorEmailList.Add(DateTime.Now + "| " + labelName + " is missing.");
                        mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + labelName + " is missing from " + passedFile + "... \r\n", Color.Red, FontStyle.Regular); }));
                    }

                    foreach (string s in labelFiles)
                    {
                        dropboxFileName.Add(s.Trim());
                        if (Path.GetFileName(s) == Path.GetFileName(previousFile))
                        {
                            errorCheck = true;
                            using (System.IO.StreamWriter errorFile = new System.IO.StreamWriter(Settings.Default.shawErrorFolder + "\\" + workingFile + ".txt", true))
                                errorFile.WriteLine(DateTime.Now + "| " + labelName + " is duplicated on the server.");
                            errorEmailList.Add(DateTime.Now + "| " + labelName + " is duplicated on the server.");
                            mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + Path.GetFileName(s) + " has a duplicate file in " + passedFile + "... \r\n", Color.Red, FontStyle.Regular); }));
                        }
                        previousFile = s;
                        dropboxQty.Add(labelQty);
                    }
                }

                if (errorCheck == true)
                {
                    if (File.Exists(Settings.Default.shawErrorFolder + "\\" + workingFile) && File.Exists(Settings.Default.shawHotfolder + "\\" + workingFile))
                    {
                        System.IO.File.Delete(Settings.Default.shawErrorFolder + "\\" + workingFile);
                    }
                    if (File.Exists(Settings.Default.shawHotfolder + "\\" + workingFile))
                    {
                        System.IO.File.Move(Settings.Default.shawHotfolder + "\\" + workingFile, Settings.Default.shawErrorFolder + "\\" + workingFile);
                    }
                    if (Settings.Default.sendEmails == true)
                    {
                        mail.SendMailShawTeamNew(workingFile, errorEmailList, true);
                        errorEmailList.Clear();
                    }
                }
                else
                {
                    List<string> indigo5600List = new List<string>();
                    List<int> indigo5600ListTicket = new List<int>();
                    List<int> hpListTicket = new List<int>();
                    List<string> cleanup6800 = Indigo6800(dropboxFileName, dropboxQty, diffPerSheet, workingFile);

                    for (int i = 0; i < dropboxFileName.Count; i++)
                    {                        
                        if (!cleanup6800.Contains(dropboxFileName[i]))
                        {
                            indigo5600List.Add(dropboxFileName[i]);
                        }
                    }                

                    List<string> list12x18;
                    List<string> list13x19;
                    List<string> hpList;
                    List<string> numberupList12x18;
                    List<string> numberupList13x19;

                    Indigo5600(indigo5600List, out list12x18, out numberupList12x18, out list13x19, out numberupList13x19, out hpList);

                    if (list12x18.Any() || list13x19.Any() || hpList.Any())
                    {
                        int fileProgressStep = (int)Math.Ceiling(((double)100) / (list12x18.Count + list13x19.Count + hpList.Count));
                        mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.Step = fileProgressStep; }));
                        if (list12x18.Any())
                        {
                            while (list12x18.Count > 0)
                            {
                                string formattedSize = "";
                                List<string> tempList12x18 = new List<string>();
                                List<string> tempNumberUpList12x18 = new List<string>();
                                List<int> dropList = new List<int>();
                                int z = 0;

                                formattedSize = pdfProcessing.FormatGetSize(list12x18[0], "trim");
                                foreach (string s in list12x18)
                                {
                                    if (pdfProcessing.FormatGetSize(s, "trim") == formattedSize)
                                    {
                                        tempList12x18.Add(s);
                                        tempNumberUpList12x18.Add(numberupList12x18[z]);
                                        dropList.Add(z);
                                        mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                                    }
                                    z++;
                                }
                                for (int i = 0; i < dropboxFileName.Count; i++)
                                {
                                    if (tempList12x18.Contains(dropboxFileName[i]))
                                    {
                                        indigo5600ListTicket.Add(i + 1);
                                    }
                                }
                                preflightPdf.PreflightPdfLayoutCombined(mainForm, tempList12x18.ToArray(), Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 12x18 Stock - Printable.pdf", 12, 18, "Indigo");
                                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 12x18 Stock - Printable.pdf", Settings.Default.shawIndigo5600 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 12x18 Stock - Printable.pdf", true);
                                if ((Settings.Default.debugOn == false))
                                {
                                    ticket.shawPrintableTicket(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 12x18 Stock", tempNumberUpList12x18, indigo5600ListTicket, formattedSize);
                                }
                                indigo5600ListTicket.Clear();
                                tempList12x18.Clear();
                                dropList.Reverse();
                                foreach (int i in dropList)
                                {
                                    list12x18.RemoveAt(i);
                                    numberupList12x18.RemoveAt(i);
                                }
                                dropList.Clear();
                                tempNumberUpList12x18.Clear();
                            }
                        }
                        if (list13x19.Any())
                        {
                            while (list13x19.Count > 0)
                            {
                                string formattedSize = "";
                                List<string> tempList13x19 = new List<string>();
                                List<string> tempNumberUpList13x19 = new List<string>();
                                List<int> dropList = new List<int>();
                                int z = 0;

                                formattedSize = pdfProcessing.FormatGetSize(list13x19[0], "trim");
                                foreach (string s in list13x19)
                                {

                                    if (pdfProcessing.FormatGetSize(s, "trim") == formattedSize)
                                    {
                                        tempList13x19.Add(s);
                                        tempNumberUpList13x19.Add(numberupList13x19[z]);
                                        dropList.Add(z);
                                        mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                                    }
                                    z++;
                                }
                                for (int i = 0; i < dropboxFileName.Count; i++)
                                {
                                    if (tempList13x19.Contains(dropboxFileName[i]))
                                    {
                                        indigo5600ListTicket.Add(i + 1);
                                    }
                                }
                                preflightPdf.PreflightPdfLayoutCombined(mainForm, tempList13x19.ToArray(), Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 13x19 Stock - Printable.pdf", 13, 19, "Indigo");
                                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 13x19 Stock - Printable.pdf", Settings.Default.shawIndigo5600 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 13x19 Stock - Printable.pdf", true);
                                if ((Settings.Default.debugOn == false))
                                {
                                    ticket.shawPrintableTicket(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 13x19 Stock", tempNumberUpList13x19, indigo5600ListTicket, formattedSize);
                                }
                                indigo5600ListTicket.Clear();
                                tempList13x19.Clear();
                                dropList.Reverse();
                                foreach (int i in dropList)
                                {
                                    list13x19.RemoveAt(i);
                                    numberupList13x19.RemoveAt(i);
                                }
                                dropList.Clear();
                                tempNumberUpList13x19.Clear();
                            }
                        }
                        if (hpList.Any())
                        {
                            while (hpList.Count > 0)
                            {                                
                                List<string> tempListHp = new List<string>();
                                List<string> tempNumberUpListHp = new List<string>();
                                List<int> dropList = new List<int>();
                                int z = 0;

                                string formattedSizeMedia = pdfProcessing.FormatGetSize(hpList[0], "media");

                                string size = SizeCheck(Path.GetFileName(hpList[0]));

                                foreach (string s in hpList)
                                {
                                    if (pdfProcessing.FormatGetSize(s, "media") == formattedSizeMedia)
                                    {
                                        tempListHp.Add(s);
                                        dropList.Add(z);
                                        mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                                    }
                                    z++;
                                }
                                for (int i = 0; i < dropboxFileName.Count; i++)
                                {
                                    if (tempListHp.Contains(dropboxFileName[i]))
                                    {
                                        hpListTicket.Add(i + 1);
                                    }
                                }
                                outputHPShaw.HPShaw(tempListHp);
                                if ((Settings.Default.debugOn == false))
                                {
                                    ticket.shawPrintableTicket(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - " + size + " - HP", tempNumberUpListHp, hpListTicket, size);
                                }
                                hpListTicket.Clear();
                                tempListHp.Clear();
                                dropList.Reverse();
                                foreach (int i in dropList)
                                {
                                    hpList.RemoveAt(i);
                                }
                                dropList.Clear();
                                tempNumberUpListHp.Clear();
                            }
                        }
                    }

                    if (Settings.Default.sendEmails == true)
                    {
                        mail.SendMailShawTeam(workingFile, false);
                    }
                    if (!workingFile.ToLower().Contains("pdf"))
                    {
                        if (File.Exists(Settings.Default.shawArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + workingFile))
                        {
                            System.IO.File.Delete(Settings.Default.shawArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + workingFile);
                        }
                        System.IO.Directory.CreateDirectory(Settings.Default.shawArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\");
                        if (File.Exists(Settings.Default.shawHotfolder + workingFile))
                        {
                            System.IO.File.Copy(Settings.Default.shawHotfolder + workingFile, Settings.Default.shawArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + workingFile);
                        }
                    }
                    if (File.Exists(Settings.Default.shawHotfolder + "\\" + workingFile))
                    {                        
                        for (int i = 1; i < validCellsCheck; i++)
                        {
                            string customer = "shaw";
                            string orderName = Path.GetFileNameWithoutExtension(workingFile);
                            string fileName = sheet.Cells[i, 0].StringValue;
                            string partNumber = sheet.Cells[i, 3].StringValue;
                            string size = SizeCheck(fileName.Trim() + ".pdf");
                            string qty = sheet.Cells[i, 5].StringValue;
                            string woNumber = sheet.Cells[i, 6].StringValue;
                            string soNumber = sheet.Cells[i, 7].StringValue;
                            string specs = sheet.Cells[i, 8].StringValue;
                            methodsMySQL.InsertOrders(customer, orderName, fileName, partNumber, size, qty, woNumber, soNumber, specs);
                        }
                        System.IO.File.Delete(Settings.Default.shawHotfolder + "\\" + workingFile);
                    }
                    exceptionErrorList.Remove(workingFile);
                }
                dropboxFileName.Clear();
                dropboxQty.Clear();
            }
            catch (Exception ex)
            {
                foreach (string e in exceptionErrorList)
                {
                    using (System.IO.StreamWriter errorFile = new System.IO.StreamWriter(Settings.Default.shawErrorFolder + "\\" + e + ".txt", true))
                    {
                        errorFile.WriteLine(DateTime.Now + "| " + ex.Message + "\r\nCheck your spreadsheet format. \r\nAsk Matt with questions.");
                        Console.WriteLine(DateTime.Now + "| " + ex.Message + "\r\nCheck your spreadsheet format. \r\nAsk Matt with questions.");
                        mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + ex.Message + "\r\nCheck your spreadsheet format. \r\nAsk Matt with questions.\r\n", Color.Red, FontStyle.Regular); }));
                        errorEmailList.Add(DateTime.Now + "| " + ex.Message + "\r\nCheck your spreadsheet format. \r\nAsk Matt with questions.");
                        if (Settings.Default.sendEmails == true)
                        {
                            mail.SendMailShawTeamNew(e, errorEmailList, true);
                            errorEmailList.Clear();
                        }
                    }
                    if (File.Exists(Settings.Default.shawErrorFolder + "\\" + e) && File.Exists(Settings.Default.shawHotfolder + "\\" + e))
                    {
                        System.IO.File.Delete(Settings.Default.shawErrorFolder + "\\" + e);
                    }
                    if (File.Exists(Settings.Default.shawHotfolder + "\\" + e))
                    {
                        System.IO.File.Move(Settings.Default.shawHotfolder + "\\" + e, Settings.Default.shawErrorFolder + "\\" + e);
                    }
                }
            }
            finally
            {
                exceptionErrorList.Clear();
            }

        }

        private string SizeCheck(string fileName)
        {
            string[] sizeCheck = Directory.GetFiles(Settings.Default.shawPdfs, fileName, SearchOption.AllDirectories);
            string formattedSizeMedia = pdfProcessing.FormatGetSize(sizeCheck[0], "media");
            string formattedSizeTrim = pdfProcessing.FormatGetSize(sizeCheck[0], "trim");
            var sizes = formattedSizeTrim.Split('x');
            double width = double.Parse(sizes[0]);
            double height = double.Parse(sizes[1]);
            if (formattedSizeMedia == formattedSizeTrim)
            {
                width = width - 3;
                height = height - 3;
            }
            string size = width.ToString() + " x " + height.ToString();
            return size;
        }

        private List<string> Indigo6800(List<string> dropboxFileName, List<int> dropboxQty, List<string> diffPerSheet, string workingFile)
        {
            List<string> removeList = new List<string>();
            List<string> tempFileNames = new List<string>();
            List<int> tempQtys = new List<int>();
            List<int> tempLines = new List<int>();

            ////////////////////////////
            string longorshort = "";
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("2 x 0.5") && !dropboxFileName[i].Contains("Generic") && (dropboxQty[i] <= 200))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                    longorshort = "Short";
                }
                else if (dropboxFileName[i].Contains("2 x 0.5") && (dropboxFileName[i].Contains("Generic") || (dropboxQty[i] >= 200)))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                    longorshort = "Long";
                }
            }
            if (tempFileNames.Any())
            {
                if (longorshort == "Short")
                {
                    diffPerSheet = outputIndigo6800Shaw.pdf2x0_5_Short(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                }
                else
                {
                    diffPerSheet = outputIndigo6800Shaw.pdf2x0_5_Long(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                }
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2x0.5 " + longorshort + ".pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2x0.5 " + longorshort + ".pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2x0.5 " + longorshort, diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            
            //////////////////////////// 

            /*///////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("2 x 0.5") && (dropboxFileName[i].Contains("Generic") || (dropboxQty[i] >= 200)))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf2x0_5_Long(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2x0.5 Long.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2x0.5 Long.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2x0.5 Long", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0]));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();

            *////////////////////////////
            

            ////////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("1.5 x 0.375"))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf1_5x0_375_6800(workingFile, tempFileNames.ToArray(), dropboxQty.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 1.5x0.375.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 1.5x0.375.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 1.5x0.375", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            

            ////////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("2 x 1") && !dropboxFileName[i].Contains("12 x 16") && !dropboxFileName[i].Contains("12 x 17"))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf2x1_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2x1.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2x1.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2x1", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////


            ////////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("3 x 1") && !dropboxFileName[i].Contains("3 x 1.5"))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf3x1_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3x1.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3x1.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3x1", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            ///////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("2 x 2") && !dropboxFileName[i].Contains("Circle"))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf2x2_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2x2.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2x2.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2x2", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            ///////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("2 x 2") && dropboxFileName[i].Contains("Circle"))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf2x2Circle_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2x2 Circle.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2x2 Circle.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2x2 Circle", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            ///////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("2.625 x 1.0625"))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf2_625x1_0625_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2.625x1.0625.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2.625x1.0625.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2.625x1.0625", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            ///////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("2.625 x 1.125"))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf2_625x1_125_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2.625x1.125.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2.625x1.125.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2.625x1.125", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            ///////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("3 x 2"))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf3x2_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3x2.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3x2.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3x2", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            ///////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("3 x 0.5"))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf3x_5_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3x0.5.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3x0.5.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3x0.5", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            ///////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("3.5 x 1.25"))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf3_5x1_25_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3.5x1.25.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3.5x1.25.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3.5x1.25", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            ///////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("3.5 x 3.5"))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf3_5x3_5_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3.5x3.5.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3.5x3.5.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3.5x3.5", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            ///////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("3.5 Triangle"))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf3_5_Triangle_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3.5 Triangle.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3.5 Triangle.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3.5 Triangle", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            ///////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if ((dropboxFileName[i].Contains("Clear")) && (dropboxFileName[i].Contains("3.25 x 1.75")))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf3_25x1_75_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3.25x1.75 Clear.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3.25x1.75 Clear.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 3.25x1.75 Clear", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            ///////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("4.3125 x 4.3125"))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf4_3125x4_3125_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 4.3125x4.3125.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 4.3125x4.3125.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 4.3125x4.3125", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            ///////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("4.25 x 1"))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf4_25x1_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 4.25x1.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 4.25x1.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 4.25x1", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            ///////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("0.5") && (dropboxFileName[i].Contains("Circle")))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf0_5x0_5_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 0.5 Circle.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 0.5 Circle.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 0.5 Circle.pdf", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            ///////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("1.5") && (dropboxFileName[i].Contains("Circle")))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf1_5x1_5_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 1.5 Circle.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 1.5 Circle.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 1.5 Circle", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            ///////////////////////////
            for (int i = 0; i < dropboxFileName.Count; i++)
            {
                if (dropboxFileName[i].Contains("2.75 x 0.312"))
                {
                    tempFileNames.Add(dropboxFileName[i]);
                    tempQtys.Add(dropboxQty[i]);
                    tempLines.Add(i + 1);
                    removeList.Add(dropboxFileName[i]);
                }
            }
            if (tempFileNames.Any())
            {
                diffPerSheet = outputIndigo6800Shaw.pdf2_75x0_312_6800(workingFile, tempFileNames.ToArray(), tempQtys.ToArray());
                System.IO.File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2.75x0.312.pdf", Settings.Default.shawIndigo6800 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2.75x0.312.pdf", true);
                if (Settings.Default.debugOn == false)
                {
                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - Printable 2.75x0.312", diffPerSheet, tempLines, pdfProcessing.FormatGetSize(tempFileNames[0], "trim"));
                }
            }
            diffPerSheet.Clear();
            tempFileNames.Clear();
            tempQtys.Clear();
            tempLines.Clear();
            ////////////////////////////

            return removeList;
        }

        private void Indigo5600(List<string> dropboxFileName, out List<string> list12x18, out List<string> numberupList12x18, out List<string> list13x19, out List<string> numberupList13x19, out List<string> hpList)
        {
            list12x18 = new List<string>();
            list13x19 = new List<string>();
            numberupList12x18 = new List<string>();
            numberupList13x19 = new List<string>();
            List<float> parseSizes = new List<float>();
            List<float> parseSizes13x19 = new List<float>();
            List<int> parseCalculate = new List<int>();
            List<int> parseCalculate13x19 = new List<int>();
            hpList = new List<string>();

            float bleed;

            foreach (string s in dropboxFileName)
            {
                bleed = pdfProcessing.PdfResize(s);
                parseSizes = pdfProcessing.GetSize(Settings.Default.tempDir + "\\" + Path.GetFileName(s), bleed);
                parseSizes.RemoveAt(0);
                parseCalculate = pdfProcessing.Calculate((12 * 72).ToString(), (18 * 72).ToString(), (parseSizes[0] - bleed).ToString(), (parseSizes[1] - bleed).ToString(), parseSizes[2].ToString());

                if (parseCalculate[0] <= 0)
                {
                    bleed = pdfProcessing.PdfResize(s);
                    parseSizes13x19 = pdfProcessing.GetSize(Settings.Default.tempDir + "\\" + Path.GetFileName(s), bleed);
                    parseSizes13x19.RemoveAt(0);
                    parseCalculate13x19 = pdfProcessing.Calculate((13 * 72).ToString(), (19 * 72).ToString(), (parseSizes13x19[0] - bleed).ToString(), (parseSizes13x19[1] - bleed).ToString(), parseSizes13x19[2].ToString());

                    if (parseCalculate13x19[0] <= 0)
                    {
                        //outputHPShaw.HPShaw(s);
                        hpList.Add(s);
                        //throw new Exception(s + " will not fit on 12x18 or 13x19 stock.  Please remove it from the spreadsheet and resubmit");
                    }
                    else
                    {
                        numberupList13x19.Add(parseCalculate13x19[0].ToString() + "up - ");
                        list13x19.Add(s);
                    }
                }
                else
                {
                    numberupList12x18.Add(parseCalculate[0].ToString() + "up - ");
                    list12x18.Add(s);
                }
            }

        }
    }
}
