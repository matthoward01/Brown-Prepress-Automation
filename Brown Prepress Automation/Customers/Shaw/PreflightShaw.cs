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
                int validCellsCheck = methods.countValidCells(Settings.Default.shawHotfolder + "\\" + workingFile, 1, 0, 0);                
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
                    if (coverFiles.Length < 1)
                    {
                        coverFiles = Directory.GetFiles(Settings.Default.shawPdfsOld, coverName, SearchOption.AllDirectories);
                    }

                    string[] linerFiles = Directory.GetFiles(Settings.Default.shawPdfs, linerName, SearchOption.AllDirectories);
                    if (linerFiles.Length < 1)
                    {
                        linerFiles = Directory.GetFiles(Settings.Default.shawPdfsOld, linerName, SearchOption.AllDirectories);
                    }

                    string[] flFiles = Directory.GetFiles(Settings.Default.shawPdfs, flName, SearchOption.AllDirectories);
                    if (flFiles.Length < 1)
                    {
                        flFiles = Directory.GetFiles(Settings.Default.shawPdfsOld, flName, SearchOption.AllDirectories);
                    }

                    string[] blFiles = Directory.GetFiles(Settings.Default.shawPdfs, blName, SearchOption.AllDirectories);
                    if (blFiles.Length < 1)
                    {
                        blFiles = Directory.GetFiles(Settings.Default.shawPdfsOld, blName, SearchOption.AllDirectories);
                    }

                    if (coverFiles == null || coverFiles.Length < 1)
                    {
                        errorCheck = true;
                        using (StreamWriter errorFile = new StreamWriter(Settings.Default.shawErrorFolder + "\\" + workingFile + ".txt", true))
                            errorFile.WriteLine(DateTime.Now + "| " + coverName + " is missing. Please Fix and resubmit spreadsheet via the hotfolder.");
                        mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + coverName + " is missing from " + passedFile + "... \r\n", Color.Red, FontStyle.Regular); }));
                    }
                    if (linerFiles == null || linerFiles.Length < 1)
                    {
                        errorCheck = true;
                        using (StreamWriter errorFile = new StreamWriter(Settings.Default.shawErrorFolder + "\\" + workingFile + ".txt", true))
                            errorFile.WriteLine(DateTime.Now + "| " + linerName + " is missing. Please Fix and resubmit spreadsheet via the hotfolder.");
                        mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + linerName + " is missing from " + passedFile + "... \r\n", Color.Red, FontStyle.Regular); }));
                    }
                    if (flFiles == null || flFiles.Length < 1)
                    {
                        errorCheck = true;
                        using (StreamWriter errorFile = new StreamWriter(Settings.Default.shawErrorFolder + "\\" + workingFile + ".txt", true))
                            errorFile.WriteLine(DateTime.Now + "| " + flName + " is missing. Please Fix and resubmit spreadsheet via the hotfolder.");
                        mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + flName + " is missing from " + passedFile + "... \r\n", Color.Red, FontStyle.Regular); }));
                    }
                    if (blFiles == null || blFiles.Length < 1)
                    {
                        errorCheck = true;
                        using (StreamWriter errorFile = new StreamWriter(Settings.Default.shawErrorFolder + "\\" + workingFile + ".txt", true))
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
                        File.Delete(Settings.Default.shawErrorFolder + "\\" + workingFile);
                    }
                    if (File.Exists(Settings.Default.shawHotfolder + "\\" + workingFile))
                    {
                        File.Move(Settings.Default.shawHotfolder + "\\" + workingFile, Settings.Default.shawErrorFolder + "\\" + workingFile);
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
                            Directory.CreateDirectory(Settings.Default.parts + "\\" + productType + "\\" + partNumberList[0].Substring(0, 8).Replace(productType + "-", "") + "000" + "\\" + partNumberList[0] + "\\");
                            File.Copy(coverList[0], Settings.Default.parts + "\\" + productType + "\\" + partNumberList[0].Substring(0, 8).Replace(productType + "-", "") + "000" + "\\" + partNumberList[0] + "\\" + partNumberList[0] + " Cover.pdf", true);
                            File.Copy(linerList[0], Settings.Default.parts + "\\" + productType + "\\" + partNumberList[0].Substring(0, 8).Replace(productType + "-", "") + "000" + "\\" + partNumberList[0] + "\\" + partNumberList[0] + " Liner.pdf", true);
                            string[] coverLinerArray = outputHPShaw.pdf1upBoard(partNumberList[0], coverList[0], linerList[0], flList[0], blList[0]);
                            File.Copy(Settings.Default.tempDir + "\\" + partNumberList[0] + " HP 40x56 - Printable.pdf", Settings.Default.parts + "\\" + productType + "\\" + partNumberList[0].Substring(0, 8).Replace(productType + "-", "") + "000" + "\\" + partNumberList[0] + "\\" + partNumberList[0] + " HP 40x56 - Printable.pdf", true);
                            coverLinerList = coverLinerArray.ToList();
                            if (Settings.Default.debugOn != true)
                            {
                                if (!board.Contains("Supplied Boards"))
                                {
                                    File.Copy(Settings.Default.tempDir + "\\" + partNumberList[0] + " HP 40x56 - Printable.pdf", Settings.Default.bluelineOutput + "\\" + partNumberList[0] + " HP 40x56 - Printable.pdf", true);
                                }
                                else
                                {
                                    foreach (string f in coverLinerList)
                                    {
                                        File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(f) + ".pdf", Settings.Default.shawHotfolder + "\\" + Path.GetFileNameWithoutExtension(f) + ".pdf", true);
                                    }
                                }
                                File.Copy(Settings.Default.tempDir + "\\" + partNumberList[0] + " HP 40x56 - Printable.pdf", Settings.Default.shawHpOutput + "\\" + partNumberList[0] + " HP 40x56 - Printable.pdf", true);

                                methodsMySQL.InsertPrepressLogAutomation(partNumberList[0]);
                            }

                            foreach (string f in coverLinerList)
                            {
                                File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(f) + ".pdf", Settings.Default.parts + "\\" + productType + "\\" + partNumberList[0].Substring(0, 8).Replace(productType + "-", "") + "000" + "\\" + partNumberList[0] + "\\" + Path.GetFileNameWithoutExtension(f) + ".pdf", true);
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
                            Directory.CreateDirectory(Settings.Default.parts + "\\" + productType + "\\" + partNumberList[0].Substring(0, 8).Replace(productType + "-", "") + "000" + "\\" + partNumberList[0] + "\\");
                            File.Copy(coverList[0], Settings.Default.parts + "\\" + productType + "\\" + partNumberList[0].Substring(0, 8).Replace(productType + "-", "") + "000" + "\\" + partNumberList[0] + "\\" + partNumberList[0] + " Cover.pdf", true);
                            File.Copy(linerList[0], Settings.Default.parts + "\\" + productType + "\\" + partNumberList[0].Substring(0, 8).Replace(productType + "-", "") + "000" + "\\" + partNumberList[0] + "\\" + partNumberList[0] + " Liner.pdf", true);
                            if (Settings.Default.debugOn == false)
                            {
                                File.Copy(coverList[0], Settings.Default.bluelineOutput + "\\" + partNumberList[0] + " Cover.pdf", true);
                                File.Copy(linerList[0], Settings.Default.bluelineOutput + "\\" + partNumberList[0] + " Liner.pdf", true);
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
                        File.Delete(Settings.Default.shawArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + workingFile);
                    }
                    Directory.CreateDirectory(Settings.Default.shawArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\");
                    if (File.Exists(Settings.Default.shawHotfolder + workingFile))
                    {
                        File.Copy(Settings.Default.shawHotfolder + workingFile, Settings.Default.shawArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + workingFile);
                    }
                }
                if (File.Exists(Settings.Default.shawHotfolder + "\\" + workingFile))
                {
                    for (int i = 1; i < validCellsCheck; i++)
                    {
                        string customer = "shaw";
                        string orderName = Path.GetFileNameWithoutExtension(workingFile);
                        string fileName = sheet.Cells[i, 1].StringValue + " - " + sheet.Cells[i, 2].StringValue;
                        if (sheet.Cells[i, 3].StringValue.Trim() != "")
                        {
                            fileName += " - " + sheet.Cells[i, 3].StringValue;
                        }
                        if (sheet.Cells[i, 4].StringValue.Trim() != "")
                        {
                            fileName += " - " + sheet.Cells[i, 4].StringValue;
                        }
                        string partNumber = sheet.Cells[i, 0].StringValue;
                        string size = sheet.Cells[i, 5].StringValue;
                        string qty = sheet.Cells[i, 8].StringValue;
                        string woNumber = "";
                        string soNumber = "";
                        string specs = sheet.Cells[i, 6].StringValue + " " + sheet.Cells[i, 7].StringValue;
                        methodsMySQL.InsertOrders(customer, orderName, fileName, partNumber, size, qty, woNumber, soNumber, specs);
                    }
                    File.Delete(Settings.Default.shawHotfolder + "\\" + workingFile);
                }
                exceptionErrorList.Remove(workingFile);
            }
            catch (Exception ex)
            {
                foreach (string e in exceptionErrorList)
                {
                    using (StreamWriter errorFile = new StreamWriter(Settings.Default.shawErrorFolder + "\\" + e + ".txt", true))
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
                        File.Delete(Settings.Default.shawErrorFolder + "\\" + e);
                    }
                    if (File.Exists(Settings.Default.shawHotfolder + "\\" + e))
                    {
                        File.Move(Settings.Default.shawHotfolder + "\\" + e, Settings.Default.shawErrorFolder + "\\" + e);
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
                List<string> dropboxSpecs = new List<string>();
                List<string> dropboxPartNumber = new List<string>();
                List<string> diffPerSheet = new List<string>();
                string workingFile = passedFile;
                exceptionErrorList.Add(workingFile);
                bool errorCheck = false;
                Workbook book = Workbook.Load(Settings.Default.shawHotfolder + "\\" + workingFile);
                Worksheet sheet = book.Worksheets[0];
                int validCellsCheck = methods.countValidCells(Settings.Default.shawHotfolder + "\\" + workingFile, 1, 0, 0);
                string previousFile = "";
                for (int i = 1; i < validCellsCheck; i++)
                {
                    string labelName = sheet.Cells[i, 0].StringValue.Trim();
                    int labelQty = Convert.ToInt32(sheet.Cells[i, 5].StringValue);
                    string labelSpecs = sheet.Cells[i, 8].StringValue.Trim();
                    string labelPartNumber = sheet.Cells[i, 3].StringValue.Trim();

                    labelName = labelName.Replace("_LR", "");
                    if (!labelName.Contains(".pdf"))
                    {
                        labelName = labelName + ".pdf";
                    }

                    string[] labelFiles = Directory.GetFiles(Settings.Default.shawPdfs, labelName, SearchOption.AllDirectories);
                    if (labelFiles.Length < 1)
                    {
                        labelFiles = Directory.GetFiles(Settings.Default.shawPdfsOld, labelName, SearchOption.AllDirectories);
                    }


                    if (labelFiles == null || labelFiles.Length < 1)
                    {
                        errorCheck = true;
                        using (StreamWriter errorFile = new StreamWriter(Settings.Default.shawErrorFolder + "\\" + workingFile + ".txt", true))
                            errorFile.WriteLine(DateTime.Now + "| " + labelName + " is missing. Was a page number forgotten?  It's also possible Shaw did not name the file consistently.");
                        errorEmailList.Add(DateTime.Now + "| " + labelName + " is missing. Was a page number forgotten?  It's also possible Shaw did not name the file consistently.");
                        mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + labelName + " is missing from " + passedFile + "... \r\n", Color.Red, FontStyle.Regular); }));
                    }

                    foreach (string s in labelFiles)
                    {
                        dropboxFileName.Add(s.Trim());
                        //TODO: Maybe Fix
                        /*if (Path.GetFileName(s) == Path.GetFileName(previousFile))
                        {
                            errorCheck = true;
                            using (StreamWriter errorFile = new StreamWriter(Settings.Default.shawErrorFolder + "\\" + workingFile + ".txt", true))
                                errorFile.WriteLine(DateTime.Now + "| " + labelName + " is duplicated on the server.");
                            errorEmailList.Add(DateTime.Now + "| " + labelName + " is duplicated on the server.");
                            mainForm.BeginInvoke(new Action(() => { mainForm.rtMain.AppendText(DateTime.Now + "| " + Path.GetFileName(s) + " has a duplicate file in " + passedFile + "... \r\n", Color.Red, FontStyle.Regular); }));
                        }*/
                        previousFile = s;
                        dropboxQty.Add(labelQty);
                        dropboxSpecs.Add(labelSpecs);
                        dropboxPartNumber.Add(labelPartNumber);
                    }
                }

                if (errorCheck == true)
                {
                    if (File.Exists(Settings.Default.shawErrorFolder + "\\" + workingFile) && File.Exists(Settings.Default.shawHotfolder + "\\" + workingFile))
                    {
                        File.Delete(Settings.Default.shawErrorFolder + "\\" + workingFile);
                    }
                    if (File.Exists(Settings.Default.shawHotfolder + "\\" + workingFile))
                    {
                        File.Move(Settings.Default.shawHotfolder + "\\" + workingFile, Settings.Default.shawErrorFolder + "\\" + workingFile);
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
                    List<int> indigo5600ListQty = new List<int>();
                    List<int> indigo5600ListTicket = new List<int>();
                    List<int> hpListTicket = new List<int>();
                    /*
                    List<string> cleanup6800 = Indigo6800(dropboxFileName, dropboxQty, dropboxSpecs, diffPerSheet, workingFile);

                    for (int i = 0; i < dropboxFileName.Count; i++)
                    {                        
                        if (!cleanup6800.Contains(dropboxFileName[i]))
                        {
                            indigo5600List.Add(dropboxFileName[i]);
                        }
                    }
                    */
                    for (int i = 0; i < dropboxFileName.Count; i++)
                    {
                        if (dropboxSpecs[i].ToLower().Contains("tray"))
                        {
                            mail.SendMailShawTray("Run tray on " + dropboxPartNumber[i] + ".");
                        }
                    }

                    List<string> list6800 = new List<string>();
                    List<int> list6800Qty = new List<int>();
                    List<int> list6800Ticket = new List<int>();
                    List<string> list12x18;
                    List<int> list12x18Qty;
                    List<string> list13x19;
                    List<int> list13x19Qty;
                    List<string> hpList;
                    List<int> hpListQty;
          
                    List<string> numberupList12x18;
                    List<string> numberupList13x19;

                    int qty6800 = 0;
                    foreach (string s in dropboxFileName)
                    {
                        if (
                            dropboxSpecs[qty6800].ToLower().Contains("mactac") ||
                            dropboxSpecs[qty6800].ToLower().Contains("tekra") ||
                            dropboxSpecs[qty6800].ToLower().Contains("60# white semi gloss with permanent adhesive") ||
                            dropboxSpecs[qty6800].ToLower().Contains("omega") ||
                            //dropboxSpecs[qty6800].ToLower().Contains("tire grip") ||
                            dropboxSpecs[qty6800].ToLower().Contains("rolls")
                           )
                        {
                            list6800.Add(s);
                            list6800Qty.Add(dropboxQty[qty6800]);
                            list6800Ticket.Add(qty6800);
                        }
                        qty6800++;
                    }
                    for (int i = 0; i < dropboxFileName.Count; i++)
                    {
                        if (!list6800.Contains(dropboxFileName[i]))
                        {
                            indigo5600List.Add(dropboxFileName[i]);
                            indigo5600ListQty.Add(dropboxQty[i]);
                        }
                    }
                    Indigo5600(indigo5600List, indigo5600ListQty, out list12x18, out list12x18Qty, out numberupList12x18, out list13x19, out list13x19Qty, out numberupList13x19, out hpList, out hpListQty);

                    if (list12x18.Any() || list13x19.Any() || hpList.Any() || list6800.Any())
                    {
                        int fileProgressStep = (int)Math.Ceiling(((double)100) / (list12x18.Count + list13x19.Count + hpList.Count + list6800.Count));
                        mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.Step = fileProgressStep; }));
                        if (list6800.Any())
                        {
                            while (list6800.Count > 0)
                            {
                                string formattedSize = "";
                                List<string> tempList6800 = new List<string>();
                                List<int> tempList6800Qty = new List<int>();
                                List<string> tempNumberUpList6800 = new List<string>();
                                List<int> dropList = new List<int>();
                                int z = 0;

                                formattedSize = pdfProcessing.FormatGetSize(list6800[0], "trim", 1);
                                foreach (string s in list6800)
                                {
                                    if (pdfProcessing.FormatGetSize(s, "trim", 1) == formattedSize)
                                    {
                                        tempList6800.Add(s);
                                        tempList6800Qty.Add(list6800Qty[z]);
                                        //tempNumberUpList6800.Add(numberupList12x18[z]);
                                        dropList.Add(z);
                                        mainForm.BeginInvoke(new Action(() => { mainForm.pbIndividual.PerformStep(); }));
                                    }
                                    z++;
                                }
                                for (int i = 0; i < dropboxFileName.Count; i++)
                                {
                                    if (tempList6800.Contains(dropboxFileName[i]))
                                    {
                                        indigo5600ListTicket.Add(i + 1);
                                    }
                                }
                                string passedFilename = "";
                                if (formattedSize == "3.25 x 1.75")
                                {
                                    passedFilename = Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " Clear.pdf";
                                    diffPerSheet =  outputIndigo6800Shaw.pdf3_25x1_75_6800(mainForm, passedFilename, tempList6800.ToArray(), tempList6800Qty.ToArray());
                                }
                                else if (formattedSize == "0.50 x 0.50")
                                {
                                    passedFilename = Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " Circle.pdf";
                                    diffPerSheet = outputIndigo6800Shaw.pdf0_5x0_5_6800(mainForm, passedFilename, tempList6800.ToArray(), tempList6800Qty.ToArray());
                                }
                                else if ((formattedSize == "2.00 x 0.50"))
                                {
                                    passedFilename = Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " Long.pdf";
                                    diffPerSheet = outputIndigo6800Shaw.pdf2x0_5_Long(mainForm, passedFilename, tempList6800.ToArray(), tempList6800Qty.ToArray());
                                }
                                else if (formattedSize == "2.00 x 1.00")
                                {
                                    passedFilename = Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + ".pdf";
                                    diffPerSheet = outputIndigo6800Shaw.pdf2x1_6800(mainForm, passedFilename, tempList6800.ToArray(), tempList6800Qty.ToArray());
                                }
                                else
                                {
                                    throw new Exception("Size is not available for web. Ask Matt to set up.");
                                }
                                //preflightPdf.PreflightPdfLayoutCombined(mainForm, tempList6800.ToArray(), Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 12x18 Stock - Printable.pdf", 12, 18, "Indigo");
                                File.Copy(passedFilename, Settings.Default.shawIndigo6800 + "\\" + Path.GetFileName(passedFilename), true);
                                if ((Settings.Default.debugOn == false))
                                {
                                    ticket.shawPrintableTicket6800(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(passedFilename), diffPerSheet, indigo5600ListTicket, formattedSize);
                                }
                                indigo5600ListTicket.Clear();
                                tempList6800.Clear();
                                dropList.Reverse();
                                foreach (int i in dropList)
                                {
                                    list6800.RemoveAt(i);
                                }
                                dropList.Clear();
                                tempNumberUpList6800.Clear();
                            }
                        }
                        if (list12x18.Any())
                        {
                            while (list12x18.Count > 0)
                            {
                                string formattedSize = "";
                                List<string> tempList12x18 = new List<string>();
                                List<int> tempList12x18Qty = new List<int>();
                                List<string> tempNumberUpList12x18 = new List<string>();
                                List<int> dropList = new List<int>();
                                int z = 0;

                                formattedSize = pdfProcessing.FormatGetSize(list12x18[0], "trim", 1);
                                foreach (string s in list12x18)
                                {
                                    if (pdfProcessing.FormatGetSize(s, "trim", 1) == formattedSize)
                                    {
                                        tempList12x18.Add(s);
                                        tempNumberUpList12x18.Add(numberupList12x18[z]);
                                        tempList12x18Qty.Add(list12x18Qty[z]);
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
                                //preflightPdf.PreflightPdfLayoutCombined(mainForm, tempList12x18.ToArray(), Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 12x18 Stock - Printable.pdf", 12, 18, "Indigo");
                                diffPerSheet = preflightPdf.PreflightPdfLayoutCombinedNew(mainForm, tempList12x18.ToArray(), Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 12x18 Stock - Printable.pdf", 12, 18, "Indigo", tempList12x18Qty.ToArray());

                                File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 12x18 Stock - Printable.pdf", Settings.Default.shawIndigo5600 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 12x18 Stock - Printable.pdf", true);
                                if ((Settings.Default.debugOn == false))
                                {
                                    ticket.shawPrintableTicket(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 12x18 Stock", tempNumberUpList12x18, indigo5600ListTicket, formattedSize, diffPerSheet, true, true);
                                }
                                indigo5600ListTicket.Clear();
                                tempList12x18.Clear();
                                tempList12x18Qty.Clear();
                                dropList.Reverse();
                                foreach (int i in dropList)
                                {
                                    list12x18.RemoveAt(i);
                                    list12x18Qty.RemoveAt(i);
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
                                List<int> tempList13x19Qty = new List<int>();
                                List<string> tempNumberUpList13x19 = new List<string>();
                                List<int> dropList = new List<int>();
                                int z = 0;

                                formattedSize = pdfProcessing.FormatGetSize(list13x19[0], "trim", 1);
                                foreach (string s in list13x19)
                                {

                                    if (pdfProcessing.FormatGetSize(s, "trim", 1) == formattedSize)
                                    {
                                        tempList13x19.Add(s);
                                        tempNumberUpList13x19.Add(numberupList13x19[z]);
                                        tempList13x19Qty.Add(z);
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
                                //preflightPdf.PreflightPdfLayoutCombined(mainForm, tempList13x19.ToArray(), Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 13x19 Stock - Printable.pdf", 13, 19, "Indigo");
                                diffPerSheet = preflightPdf.PreflightPdfLayoutCombinedNew(mainForm, tempList13x19.ToArray(), Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 13x19 Stock - Printable.pdf", 13, 19, "Indigo", tempList13x19Qty.ToArray());

                                File.Copy(Settings.Default.tempDir + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 13x19 Stock - Printable.pdf", Settings.Default.shawIndigo5600 + "\\" + Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 13x19 Stock - Printable.pdf", true);
                                if ((Settings.Default.debugOn == false))
                                {
                                    ticket.shawPrintableTicket(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - " + formattedSize + " - 13x19 Stock", tempNumberUpList13x19, indigo5600ListTicket, formattedSize, diffPerSheet, true, true);
                                }
                                indigo5600ListTicket.Clear();
                                tempList13x19.Clear();
                                tempList13x19Qty.Clear();
                                dropList.Reverse();
                                foreach (int i in dropList)
                                {
                                    list13x19.RemoveAt(i);
                                    list13x19Qty.RemoveAt(i);
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
                                List<int> tempListHpQty = new List<int>();
                                List<string> tempNumberUpListHp = new List<string>();
                                List<string> tempPartNumber = new List<string>();
                                List<int> dropList = new List<int>();
                                int z = 0;
                                string formattedSizeMedia = pdfProcessing.FormatGetSize(hpList[0], "trim", 1);

                                string size = SizeCheck(Path.GetFileName(hpList[0]));

                                foreach (string s in hpList)
                                {
                                    if (pdfProcessing.FormatGetSize(s, "trim", 1) == formattedSizeMedia)
                                    {
                                        tempListHp.Add(s);
                                        tempListHpQty.Add(z);
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
                                        tempPartNumber.Add(dropboxPartNumber[i]);
                                        methodsMySQL.InsertPrepressLogAutomation(dropboxPartNumber[i]);
                                    }
                                }
                                outputHPShaw.HPShaw(tempListHp, tempPartNumber, mainForm);                                
                                if ((Settings.Default.debugOn == false))
                                {
                                    ticket.shawPrintableTicket(Settings.Default.shawHotfolder + "\\" + workingFile, Path.GetFileNameWithoutExtension(workingFile) + " - " + size + " - HP", tempNumberUpListHp, hpListTicket, size, diffPerSheet, false, false);
                                }
                                hpListTicket.Clear();
                                tempListHp.Clear();
                                tempListHpQty.Clear();
                                dropList.Reverse();
                                foreach (int i in dropList)
                                {
                                    hpList.RemoveAt(i);
                                    hpListQty.RemoveAt(i);
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
                            File.Delete(Settings.Default.shawArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + workingFile);
                        }
                        Directory.CreateDirectory(Settings.Default.shawArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\");
                        if (File.Exists(Settings.Default.shawHotfolder + workingFile))
                        {
                            File.Copy(Settings.Default.shawHotfolder + workingFile, Settings.Default.shawArchiveFolder + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + workingFile);
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
                        File.Delete(Settings.Default.shawHotfolder + "\\" + workingFile);
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
                    using (StreamWriter errorFile = new StreamWriter(Settings.Default.shawErrorFolder + "\\" + e + ".txt", true))
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
                        File.Delete(Settings.Default.shawErrorFolder + "\\" + e);
                    }
                    if (File.Exists(Settings.Default.shawHotfolder + "\\" + e))
                    {
                        File.Move(Settings.Default.shawHotfolder + "\\" + e, Settings.Default.shawErrorFolder + "\\" + e);
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
            if (sizeCheck.Length < 1)
            {
                sizeCheck = Directory.GetFiles(Settings.Default.shawPdfsOld, fileName, SearchOption.AllDirectories);
            }

            string formattedSizeMedia = pdfProcessing.FormatGetSize(sizeCheck[0], "media", 1);
            string formattedSizeTrim = pdfProcessing.FormatGetSize(sizeCheck[0], "trim", 1);
            var sizes = formattedSizeTrim.Split('x');
            double width = double.Parse(sizes[0]);
            double height = double.Parse(sizes[1]);
            double area = Double.Parse(sizes[0].Trim()) * Double.Parse(sizes[1].Trim());
            if (area >= 216)
            {
                if (formattedSizeMedia == formattedSizeTrim)
                {
                    width = width - 3;
                    height = height - 3;
                }
            }
            string size = width.ToString() + " x " + height.ToString();
            return size;
        }        

        private void Indigo5600(List<string> dropboxFileName, List<int> dropboxListQty, out List<string> list12x18, out List<int> list12x18Qty, out List<string> numberupList12x18, out List<string> list13x19, out List<int> list13x19Qty, out List<string> numberupList13x19, out List<string> hpList, out List<int> hpListQty)
        {
            list12x18 = new List<string>();
            list13x19 = new List<string>();
            list12x18Qty = new List<int>();
            list13x19Qty = new List<int>();
            numberupList12x18 = new List<string>();
            numberupList13x19 = new List<string>();
            List<float> parseSizes = new List<float>();
            List<float> parseSizes13x19 = new List<float>();
            List<int> parseCalculate = new List<int>();
            List<int> parseCalculate13x19 = new List<int>();
            hpList = new List<string>();
            hpListQty = new List<int>();
            int counter = 0;
            float bleed;

            foreach (string s in dropboxFileName)
            {
                bleed = pdfProcessing.PdfResize(s);
                parseSizes = pdfProcessing.GetSize(Settings.Default.tempDir + "\\" + Path.GetFileName(s), bleed);
                parseSizes.RemoveAt(0);
                parseCalculate = pdfProcessing.Calculate((12 * 72).ToString(), (18 * 72).ToString(), (parseSizes[0] - bleed).ToString(), (parseSizes[1] - bleed).ToString(), parseSizes[2].ToString());

                if (parseCalculate[0] <= 0 || (s.ToLower().Contains("lookbook")) || (s.ToLower().Contains("photopack")) || (pdfProcessing.FormatGetSize(s, "trim", 1) == "6.00 x 12.00") || (pdfProcessing.FormatGetSize(s, "trim", 1) == "7.00 x 12.00"))
                {
                    bleed = pdfProcessing.PdfResize(s);
                    parseSizes13x19 = pdfProcessing.GetSize(Settings.Default.tempDir + "\\" + Path.GetFileName(s), bleed);
                    parseSizes13x19.RemoveAt(0);
                    parseCalculate13x19 = pdfProcessing.Calculate((13 * 72).ToString(), (19 * 72).ToString(), (parseSizes13x19[0] - bleed).ToString(), (parseSizes13x19[1] - bleed).ToString(), parseSizes13x19[2].ToString());

                    if ((parseCalculate13x19[0] <= 0) || (s.ToLower().Contains("lookbook")) || (s.ToLower().Contains("photopack")) || (pdfProcessing.FormatGetSize(s, "trim", 1) == "6.00 x 12.00") || (pdfProcessing.FormatGetSize(s, "trim", 1) == "7.00 x 12.00"))
                    {
                        //outputHPShaw.HPShaw(s);
                        hpList.Add(s);
                        hpListQty.Add(dropboxListQty[counter]);
                        //throw new Exception(s + " will not fit on 12x18 or 13x19 stock.  Please remove it from the spreadsheet and resubmit");
                    }
                    else
                    {
                        numberupList13x19.Add(parseCalculate13x19[0].ToString() + "up - ");
                        list13x19.Add(s);
                        list13x19Qty.Add(dropboxListQty[counter]);
                    }
                }
                else
                {
                    numberupList12x18.Add(parseCalculate[0].ToString() + "up - ");
                    list12x18.Add(s);
                    list12x18Qty.Add(dropboxListQty[counter]);
                }
                counter++;
            }
        }
    }
}
